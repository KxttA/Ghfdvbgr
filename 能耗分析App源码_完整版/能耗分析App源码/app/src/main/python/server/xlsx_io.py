"""
纯标准库读写 .xlsx（不依赖 openpyxl / pandas）
仅支持本系统需要的简单场景：读第一个工作表、写带表头的表格。
"""
import io
import os
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def _drop_merged_dups(cells, merges, row_no):
    """合并单元格区域里，只有左上角保留值，其余清空"""
    if not merges:
        return
    try:
        rn = int(row_no)
    except Exception:
        return
    for a, b in merges:
        ca, ra = re.match(r'([A-Z]+)(\d+)', a).groups()
        cb, rb = re.match(r'([A-Z]+)(\d+)', b).groups()
        ra, rb = int(ra), int(rb)
        if ra <= rn <= rb:
            lo, hi = _col_index(ca), _col_index(cb)
            for ci in range(lo + 1, hi + 1):
                cells.pop(ci, None)


def _col_index(ref):
    """A1 -> 1, B1 -> 2"""
    m = re.match(r'([A-Z]+)', ref or '')
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n


def read_xlsx(data, sheet=None, sheet_name=None):
    """sheet: 第几张表(从0起)；sheet_name: 按名称找。
    sheet_name 有值时会自动在全部工作表中查找含该名称的表。
    """
    z = zipfile.ZipFile(io.BytesIO(data))
    names = z.namelist()

    # 建立 sheet 名 -> 文件 的映射
    name2file = {}
    if 'xl/workbook.xml' in names and 'xl/_rels/workbook.xml.rels' in names:
        wbx = ET.fromstring(z.read('xl/workbook.xml'))
        relsx = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rid2target = {}
        for rel in relsx:
            rid2target[rel.get('Id')] = rel.get('Target')
        for sh in wbx.iter('{%s}sheet' % NS):
            rid = sh.get('{%s}id' % RNS)
            tgt = rid2target.get(rid, '')
            if tgt:
                tgt = tgt.lstrip('/')
                if not tgt.startswith('xl/'):
                    tgt = 'xl/' + tgt
                name2file[sh.get('name')] = tgt

    # 共享字符串
    shared = []
    if 'xl/sharedStrings.xml' in names:
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall('{%s}si' % NS):
            texts = [t.text or '' for t in si.iter('{%s}t' % NS)]
            shared.append(''.join(texts))

    # 选择工作表
    sheet_name_file = 'xl/worksheets/sheet1.xml'
    if sheet_name:
        hit = None
        for nm, f in name2file.items():
            if nm == sheet_name:
                hit = f; break
        if not hit:
            for nm, f in name2file.items():
                if sheet_name in nm:
                    hit = f; break
        if hit:
            sheet_name_file = hit
    elif sheet is not None and name2file:
        ordered = list(name2file.values())
        if 0 <= sheet < len(ordered):
            sheet_name_file = ordered[sheet]
    if sheet_name_file not in names:
        cands = sorted([n for n in names if n.startswith('xl/worksheets/sheet')])
        if not cands:
            return []
        sheet_name_file = cands[0]
    root = ET.fromstring(z.read(sheet_name_file))

    # 合并单元格：把起始单元格的值铺到整片区域对应的首列
    merges = []
    for mc in root.iter('{%s}mergeCell' % NS):
        rng = mc.get('ref') or ''
        if ':' in rng:
            a, b = rng.split(':')
            merges.append((a, b))

    out = []
    for row in root.iter('{%s}row' % NS):
        cells = {}
        for c in row.findall('{%s}c' % NS):
            ref = c.get('r') or ''
            idx = _col_index(ref)
            t = c.get('t')
            v = c.find('{%s}v' % NS)
            if t == 'inlineStr':
                is_ = c.find('{%s}is' % NS)
                val = ''.join(x.text or '' for x in is_.iter('{%s}t' % NS)) if is_ is not None else ''
            elif v is None:
                val = ''
            elif t == 's':
                try:
                    val = shared[int(v.text)]
                except Exception:
                    val = ''
            elif t == 'str':
                val = v.text or ''
            else:
                txt = (v.text or '').strip()
                try:
                    num = float(txt)
                    val = int(num) if num == int(num) else num
                except Exception:
                    val = txt
            cells[idx] = val
        # 合并区域内非首格的值会被 Excel 重复写入，这里清掉，避免被当成两条数据
        _drop_merged_dups(cells, merges, row.get('r'))
        if not cells:
            out.append([])
            continue
        mx = max(cells)
        out.append([cells.get(i, '') for i in range(1, mx + 1)])
    return out


def _cell_xml(col, row, value):
    ref = '%s%d' % (col, row)
    if value is None or value == '':
        return ''
    if isinstance(value, bool):
        return '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, '是' if value else '否')
    if isinstance(value, (int, float)):
        return '<c r="%s"><v>%s</v></c>' % (ref, value)
    s = escape(str(value))
    return ('<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s))


def write_xlsx(rows, sheet_title='Sheet1', widths=None):
    """rows = [[...],...]，自动套正式样式（标题/表头/边框/列宽/冻结首行）"""
    return build_workbook([dict(title=sheet_title, rows=rows, widths=widths)])


def write_xlsx_multi(sheets):
    """sheets = [(title, rows), ...] 或 [dict(title, rows, charts=...)] -> 多工作表 xlsx"""
    norm = []
    for sh in sheets:
        if isinstance(sh, dict):
            norm.append(sh)
        else:
            norm.append(dict(title=sh[0], rows=sh[1]))
    return build_workbook(norm)



# ============================================================
# 正式报表输出：标题/表头样式、边框、数字格式、列宽、冻结首行、原生图表
# ============================================================
CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
DRAW_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'

_STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="%s">'
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>'
    '<fonts count="4">'
    '<font><sz val="11"/><name val="等线"/></font>'
    '<font><b/><sz val="11"/><name val="等线"/></font>'
    '<font><b/><sz val="15"/><color rgb="FF1F4E79"/><name val="等线"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="等线"/></font>'
    '</fonts>'
    '<fills count="4">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFEDF2F8"/><bgColor indexed="64"/></patternFill></fill>'
    '</fills>'
    '<borders count="2">'
    '<border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border><left style="thin"><color rgb="FFB7C3D0"/></left><right style="thin"><color rgb="FFB7C3D0"/></right>'
    '<top style="thin"><color rgb="FFB7C3D0"/></top><bottom style="thin"><color rgb="FFB7C3D0"/></bottom><diagonal/></border>'
    '</borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="8">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'                                 # 0 default
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">'
    '<alignment horizontal="center" vertical="center"/></xf>'                                        # 1 标题
    '<xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
    '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>'                           # 2 表头
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">'
    '<alignment vertical="center" wrapText="1"/></xf>'                                               # 3 文本
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'  # 4 数字
    '<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'  # 5 分节标签
    '<xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>'  # 6 数字加粗(合计)
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>'   # 7 文本加粗
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>' % NS)


def _col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _disp_w(v):
    s = str(v if v is not None else '')
    return sum(2 if ord(ch) > 127 else 1 for ch in s)


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _styled_sheet_xml(rows, widths=None, charts_drawing_rid=None):
    """rows: list[list]；自动识别：单格行=标题(合并)、其后一行=表头、单格非空行=分节标签、合计行加粗"""
    ncols = max((len(r) for r in rows if r), default=1)
    # 列宽
    if not widths:
        widths = [0] * ncols
        for r in rows:
            for ci, v in enumerate(r or []):
                if not _is_num(v):
                    widths[ci] = max(widths[ci], min(_disp_w(v) + 4, 44))
        widths = [max(w, 9) for w in widths]
    cols_xml = '<cols>%s</cols>' % ''.join(
        '<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w)
        for i, w in enumerate(widths))

    # 行样式识别
    title_row, header_row = None, None
    for ri, r in enumerate(rows):
        vals = [v for v in (r or []) if v not in (None, '')]
        if title_row is None and len(vals) == 1 and len(rows) > ri + 1 and ri == 0:
            title_row = ri
            continue
        if title_row is not None and header_row is None and ri == title_row + 1 and vals:
            header_row = ri
            break
        if title_row is None:
            header_row = ri
            break

    body, merges = [], []
    for ri, row in enumerate(rows):
        if not row:
            continue
        vals = [v for v in row if v not in (None, '')]
        cells = []
        for ci, v in enumerate(row):
            col_l = _col_letter(ci + 1)
            ref = '%s%d' % (col_l, ri + 1)
            if v in (None, ''):
                continue
            st = 0
            if ri == title_row:
                st = 1
                if len(vals) == 1 and ncols > 1:
                    merges.append('A%d:%s%d' % (ri + 1, _col_letter(ncols), ri + 1))
            elif ri == header_row:
                st = 2
            elif len(vals) == 1:
                st = 5
                merges.append('A%d:%s%d' % (ri + 1, _col_letter(max(ncols, len(row) or 1)), ri + 1))
            elif _is_num(v):
                first = str(row[0] or '')
                st = 6 if ('合计' in first or '小计' in first) else 4
            else:
                st = 7 if ('合计' in str(row[0] or '') or '占比' in str(v)) else 3
            if _is_num(v):
                cells.append('<c r="%s" s="%d"><v>%s</v></c>' % (ref, st, v))
            else:
                cells.append('<c r="%s" t="inlineStr" s="%d"><is><t xml:space="preserve">%s</t></is></c>'
                             % (ref, st, escape(str(v))))
        body.append('<row r="%d">%s</row>' % (ri + 1, ''.join(cells)))

    pane = ('<sheetViews><sheetView workbookViewId="0"><pane ySplit="%d" topLeftCell="A%d" '
            'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            % ((header_row + 1) if header_row else 1, (header_row + 2) if header_row else 2))
    merge_xml = ('<mergeCells count="%d">%s</mergeCells>' % (len(merges), ''.join('<mergeCell ref="%s"/>' % m for m in merges))) if merges else ''
    draw_xml = '<drawing r:id="%s"/>' % charts_drawing_rid if charts_drawing_rid else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="%s" xmlns:r="%s">%s<sheetData>%s</sheetData>%s%s%s</worksheet>'
            % (NS, RNS, pane, ''.join(body), merge_xml, cols_xml, draw_xml))


def _chart_xml(title, kind, cat_ref, cat_vals, sers, colors):
    """kind: col/bar/line；sers: [(name, val_ref, [values])]"""
    dirn = 'bar' if kind == 'bar' else 'col'
    ser_xml = ''
    for i, (name, vref, vals) in enumerate(sers):
        pts_s = ''.join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (j, escape(str(v))) for j, v in enumerate(cat_vals))
        pts_n = ''.join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (j, v) for j, v in enumerate(vals))
        ser_xml += ('<c:ser><c:idx val="%d"/><c:order val="%d"/>'
                    '<c:tx><c:strRef><c:f>%s</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>'
                    '<c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr>'
                    '<c:cat><c:strRef><c:f>%s</c:f><c:strCache><c:ptCount val="%d"/>%s</c:strCache></c:strRef></c:cat>'
                    '<c:val><c:numRef><c:f>%s</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>%s</c:numCache></c:numRef></c:val>'
                    '</c:ser>' % (i, i, escape(name), escape(name), colors[i % len(colors)],
                                  escape(cat_ref), len(cat_vals), pts_s, escape(vref), len(vals), pts_n))
    tag = 'lineChart' if kind == 'line' else 'barChart'
    extra = '' if kind == 'line' else '<c:barDir val="%s"/><c:grouping val="clustered"/><c:varyColors val="0"/>' % dirn
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<c:chartSpace xmlns:c="%s" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="%s">'
            '<c:roundedCorners val="0"/><c:chart>'
            '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100"/></a:pPr><a:r><a:rPr lang="zh-CN" sz="1100"/><a:t>%s</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
            '<c:autoTitleDeleted val="0"/>'
            '<c:plotArea><c:layout/>'
            '<c:%s>%s%s'
            '<c:axId val="100"/><c:axId val="101"/></c:%s>'
            '<c:catAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="101"/></c:catAx>'
            '<c:valAx><c:axId val="101"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="100"/></c:valAx>'
            '</c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
            '</c:chart></c:chartSpace>'
            % (CHART_NS, RNS, escape(title), tag, ser_xml, extra, tag))


def _drawing_xml(anchors):
    fr = []
    for i, (c1, r1, c2, r2, cid) in enumerate(anchors):
        fr.append('<xdr:twoCellAnchor>'
                  '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                  '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                  '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
                  '<xdr:cNvPr id="%d" name="Chart %d"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
                  '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                  '<a:graphic><a:graphicData uri="%s">'
                  '<c:chart xmlns:c="%s" xmlns:r="%s" r:id="rId%d"/>'
                  '</a:graphicData></a:graphic></xdr:graphicFrame>'
                  '<xdr:clientData/></xdr:twoCellAnchor>'
                  % (c1, r1, c2, r2, i + 2, i + 1, DRAW_NS, CHART_NS, RNS, cid))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<xdr:wsDr xmlns:xdr="%s" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">%s</xdr:wsDr>'
            % (DRAW_NS, ''.join(fr)))


def build_workbook(sheets):
    """sheets: [(title, rows)] 或 [dict(title, rows, widths, charts)]
    charts: [dict(kind='col'|'bar'|'line', title, cat='A4:A15', cat_vals=[...],
                  sers=[('实收','C4:C15',[...])], anchor='G4', size=(9,15))]"""
    def col_l(n):
        s = ''
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    n = len(sheets)
    ct_over, wb_sheets, wb_rels = [], [], []
    extra_files = []      # (zip 路径, 内容)
    chart_id = 0
    colors = ['4472C4', 'ED7D31', '70AD47', 'FFC000', '5B9BD5', 'A5A5A5']
    for i, sh in enumerate(sheets, 1):
        if isinstance(sh, dict):
            title, rows = sh['title'], sh['rows']
            charts = sh.get('charts')
            widths = sh.get('widths')
        else:
            title, rows, charts, widths = sh[0], sh[1], None, None
        wb_sheets.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(title), i, i))
        ct_over.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
        wb_rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, i))
        draw_rid = None
        if charts:
            chart_id += 1
            draw_rid = 'rIdD%d' % i
            anchors = []
            chart_rids = []
            for k, ch in enumerate(charts, 1):
                chart_id += 1
                a_col = 0
                a_row = 0
                if ch.get('anchor'):
                    a_col = sum((ord(x) - 64) * 26 ** p for p, x in
                                enumerate(reversed(re.match(r'[A-Z]+', ch['anchor']).group())))
                    a_col -= 1
                    a_row = int(re.sub(r'[A-Z]', '', ch['anchor'])) - 1
                w, h = ch.get('size', (10, 15))
                anchors.append((a_col, a_row, a_col + w, a_row + h, k))
                extra_files.append(('xl/charts/chart%d.xml' % chart_id,
                                    _chart_xml(ch.get('title', ''), ch.get('kind', 'col'),
                                               ch['cat'], ch.get('cat_vals', []),
                                               ch['sers'], colors)))
                chart_rids.append((k, chart_id))
            extra_files.append(('xl/drawings/drawing%d.xml' % i, _drawing_xml(
                [(a1, b1, a2, b2, rid) for (a1, b1, a2, b2, rid), (_k, cid) in zip(anchors, chart_rids)])))
            ct_over.append('<Override PartName="/xl/drawings/drawing%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' % i)
            for k, cid in chart_rids:
                ct_over.append('<Override PartName="/xl/charts/chart%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' % cid)
            sheet_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                          '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing%d.xml"/>' % (draw_rid, i)
                          + ''.join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart%d.xml"/>' % (k, cid)
                                    for k, cid in chart_rids)
                          + '</Relationships>')
            extra_files.append(('xl/worksheets/_rels/sheet%d.xml.rels' % i, sheet_rels))
        extra_files.insert(0, ('xl/worksheets/sheet%d.xml' % i,
                               _styled_sheet_xml(rows, widths, draw_rid)))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          + ''.join(ct_over) + '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets></workbook>'
          % (NS, RNS, ''.join(wb_sheets)))
    wb_rels_x = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                 + ''.join(wb_rels) + '</Relationships>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('xl/workbook.xml', wb)
        z.writestr('xl/_rels/workbook.xml.rels', wb_rels_x)
        z.writestr('xl/styles.xml', _STYLES)
        for path, content in extra_files:
            z.writestr(path, content)
    return buf.getvalue()
