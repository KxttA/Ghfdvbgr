# -*- coding: utf-8 -*-
"""Excel 抄表数据导入：自动识别表头日期、表计类型、系统分类"""
import re, datetime, io
import openpyxl

try:
    from server.db import meter_save, reading_save, meters, conn, now
except ImportError:
    from db import meter_save, reading_save, meters, conn, now

# ---------- 系统分类规则 ----------
ELEC_RULES = [
    (r'消防|应急|排烟|送风|防火|喷淋|消控室消防|消防泵房', '消防'),
    (r'电梯|扶梯|直梯', '电梯扶梯'),
    (r'空调|冷机|新风|KTC|KTD', '空调用电'),
    (r'潜污泵|污水泵|排水泵|集水坑|污提', '排水泵'),
    (r'商铺照明', '商铺照明'),
    (r'景观照明|景观|亮化', '景观照明'),
    (r'公共照明|照明', '照明'),
    (r'安防|监控|门禁', '安防'),
    (r'临电|E区临电|F区临电', '临电转供'),
    (r'动力|排油烟|排风', '动力排风'),
]
WATER_RULES = [
    (r'卫生间|女卫|男卫|清扫间|开水间|母婴室|无障碍', '卫生间'),
    (r'雾森', '绿化雾森'),
    (r'景观|水景', '水景'),
    (r'灌溉|取水器|绿化', '绿化灌溉'),
    (r'消防', '消防补水'),
    (r'中水补水|中水', '中水补水'),
    (r'污提|隔油|设备间', '设备间'),
]

ZONE_RE = re.compile(r'([A-F])区')


def guess_sys(text, kind):
    t = text or ''
    rules = ELEC_RULES if kind == 'elec' else WATER_RULES
    for pat, name in rules:
        if re.search(pat, t, re.I):
            return name
    return '其他'


def guess_zone(*texts):
    for t in texts:
        if not t:
            continue
        m = ZONE_RE.search(str(t))
        if m:
            return m.group(1) + '区'
        if '北通道' in str(t):
            return '北通道'
        if '下沉广场' in str(t):
            return '下沉广场'
        if '地上' in str(t):
            return '地上广场'
    return ''


def parse_date(v):
    """支持 2026.3.1 / 2026-03-01 / datetime / Excel 序列号"""
    if v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    if not s:
        return None
    m = re.match(r'^(\d{4})[.\-/年](\d{1,2})[.\-/月](\d{1,2})', s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return datetime.date(y, mo, d).isoformat()
        except ValueError:
            return None
    m = re.match(r'^(\d{4})[.\-/](\d{1,2})$', s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        try:
            return datetime.date(y, mo, 1).isoformat()
        except ValueError:
            return None
    return None


def to_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '')
    if not s or s in ('-', '/', '无', '不显示', '—'):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def detect_header(ws, max_scan=8):
    """找到包含多个日期列的表头行，返回 (行号, {列号: date_str}, 前置列名映射)"""
    best = None
    for r in range(1, min(max_scan, ws.max_row) + 1):
        dates = {}
        for c in range(1, ws.max_column + 1):
            d = parse_date(ws.cell(r, c).value)
            if d:
                dates[c] = d
        if len(dates) >= 2:
            if best is None or len(dates) > len(best[1]):
                best = (r, dates)
    if not best:
        return None, {}, {}
    hr, dates = best
    cols = {}
    first_date_col = min(dates)
    for c in range(1, first_date_col):
        v = ws.cell(hr, c).value
        if v:
            cols[c] = str(v).strip()
    return hr, dates, cols


def classify_columns(cols):
    """把表头列名映射到字段"""
    m = {}
    for c, name in cols.items():
        n = name
        if '序号' in n:
            m[c] = 'seq'
        elif '配电柜' in n or '柜名' in n:
            m[c] = 'panel'
        elif '编号' in n and ('表' in n or '水表' in n):
            m[c] = 'code'
        elif '编号' in n:
            m[c] = 'code'
        elif '配电室' in n:
            m[c] = 'loc'
        elif '控制区域' in n or '用途' in n:
            m[c] = 'purpose'
        elif '位置' in n or '名称' in n:
            m[c] = 'name'
    return m


def import_sheet(ws, kind, update_existing=True, log=None):
    """
    导入一个 sheet，返回 (新增表数, 更新表数, 写入读数条数, 跳过行)
    """
    hr, dates, cols = detect_header(ws)
    if not dates:
        raise ValueError('未识别到日期列，请确认表头为 2026.3.1 这类格式')
    cmap = classify_columns(cols)

    # 表计识别：优先用 code（表编号）作为唯一键，其次 name
    existing = {}
    for m in meters(kind):
        key = (m['code'] or '').strip() or ('N:' + m['name'])
        existing[key] = m

    added = updated = reads = skipped = 0
    buf = []
    for r in range(hr + 1, ws.max_row + 1):
        vals = {f: ws.cell(r, c).value for c, f in cmap.items()}
        row_vals = {c: ws.cell(r, c).value for c in dates}
        if not any(v is not None and str(v).strip() for v in row_vals.values()):
            continue
        def _s(v):
            return '' if v is None else str(v).strip().replace('\n', ' ')
        name = _s(vals.get('name'))
        code = _s(vals.get('code'))
        purpose = _s(vals.get('purpose'))
        panel = _s(vals.get('panel'))
        loc = _s(vals.get('loc'))
        if not name and not code:
            continue
        if name in ('无', '备用', '备用3') and not code:
            skipped += 1
            continue
        if (purpose and purpose.startswith('无')) or name.startswith('无'):
            skipped += 1
            continue

        # 名称合并：位置 + 用途，便于识别
        disp = name or purpose
        if purpose and purpose not in disp and len(purpose) > 2:
            disp = disp if len(disp) > len(purpose) else purpose

        key = code or ('N:' + name + '|' + loc + '|' + purpose)
        meta = existing.get(key)
        is_total = 1 if ('总表' in name or '总抄表' in purpose or '总表' in purpose
                         or '区总水表' in name or '总水表' in name) else 0
        z = '' if is_total else guess_zone(panel, purpose, name, loc)
        sysname = guess_sys(' '.join([name, purpose, panel]), kind)
        # 水源：中水（绿化灌溉等）与自来水分开统计；“中水补水”本身接自来水
        if kind == 'water':
            source = 'tap' if (sysname == '中水补水' or '补水' in (name + purpose)) else (
                'reclaimed' if '中水' in (name + purpose) else 'tap')
            if source == 'reclaimed':
                sysname = '中水用水' if sysname in ('中水补水', '其他') else sysname
        else:
            source = 'tap'
        # 国网关口表（项目总表，配电室高压侧）
        is_grid = 1 if re.search(r'国网|关口|高压', name + ' ' + panel + ' ' + loc) else 0
        # 主备状态：名称或用途含「备用」的视为备用表，不计入任何合计
        use_status = 'backup' if ('备用' in name or '备用' in purpose) else 'main'
        # 二级总表：区域总表 / BDEF区总水表 等（隶在一级总表下，仅兜底不叠加）
        is_sub_total = 0
        if is_total and not is_grid:
            if re.search(r'区总水表|区域总表|分区总表|BDEF', name + ' ' + purpose):
                is_sub_total, is_total = 1, 0
        if is_grid:
            is_total = is_sub_total = 0

        if meta:
            if update_existing:
                meta.update(dict(name=disp, sys=sysname, zone=z or meta['zone'],
                                 loc=loc, panel=panel, purpose=purpose, is_total=is_total,
                                 is_sub_total=is_sub_total, is_grid=is_grid,
                                 use_status=use_status, source=source))
                meter_save(meta)
                updated += 1
        else:
            mid = meter_save(dict(code=code, name=disp, kind=kind, sys=sysname,
                                  zone=z, loc=loc, panel=panel, purpose=purpose,
                                  mult=1, is_total=is_total, is_sub_total=is_sub_total,
                                  enabled=1, source=source, is_grid=is_grid,
                                  use_status=use_status, factor=1))
            meta = dict(id=mid, code=code, name=disp)
            existing[key] = meta
            added += 1

        for c, dstr in dates.items():
            v = to_num(ws.cell(r, c).value)
            if v is None:
                continue
            buf.append((meta['id'], dstr, v))
            reads += 1

    # 批量写入读数（一次提交，避免长时间持锁）
    if buf:
        cx = conn()
        cx.execute('PRAGMA busy_timeout=20000')
        ts = now()
        cx.executemany(
            'INSERT INTO reading(meter_id,rdate,val,note,created) VALUES(?,?,?,?,?) '
            'ON CONFLICT(meter_id,rdate) DO UPDATE SET val=excluded.val',
            [(a, b, c, '', ts) for a, b, c in buf])
        cx.commit(); cx.close()
    return added, updated, reads, skipped


def import_workbook(data, kind='auto', update_existing=True):
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    res = []
    for ws in wb.worksheets:
        title = ws.title
        if kind == 'auto':
            k = 'water' if ('水' in title and '电' not in title) else 'elec'
            if '电' in title:
                k = 'elec'
        else:
            k = kind
        try:
            a, u, r, s = import_sheet(ws, k, update_existing)
            res.append(dict(sheet=title, kind=k, added=a, updated=u, reads=r, skipped=s, ok=True))
        except Exception as e:
            res.append(dict(sheet=title, kind=k, ok=False, error=str(e)))
    return res


# ---------- 表计台账 导入/导出 ----------
LEDGER_HEAD = ['表编号', '表名称', '类型', '系统分类', '区域', '安装位置', '配电柜/面板',
               '用途', '倍率', '计量层级', '主备状态', '折算系数',
               '水源(自来水/中水)', '国网关口表(是/否)', '启用']


def _lv_text(m):
    """计量层级文字"""
    if m.get('is_grid'):
        return '项目关口总表'
    if m.get('is_total'):
        return '一级总表'
    if m.get('is_sub_total'):
        return '二级总表'
    return '分表'


def export_ledger():
    try:
        from server.xlsx_io import write_xlsx
    except ImportError:
        from xlsx_io import write_xlsx
    data = [['表计台账'], [], LEDGER_HEAD]
    for m in meters(None):
        data.append([m['code'], m['name'], '电' if m['kind'] == 'elec' else '水',
                     m['sys'], m['zone'], m['loc'], m['panel'], m['purpose'],
                     m['mult'], _lv_text(m),
                     '备用' if m.get('use_status') == 'backup' else '主用',
                     m.get('factor') or 1,
                     '中水' if m.get('source') == 'reclaimed' else '自来水',
                     '是' if m.get('is_grid') else '',
                     '' if m.get('enabled') else '停用'])
    return write_xlsx(data, '台账', [12, 26, 6, 12, 10, 18, 16, 18, 8, 13, 9, 9, 14, 13, 8])


def import_ledger(data, update_existing=True):
    """导入表计台账（不含读数）。返回 (新增, 更新, 跳过)"""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    added = updated = skipped = 0
    existing = {}
    for m in meters(None):
        key = (m['code'] or '').strip() or ('N:' + m['name'])
        existing[key] = m
    for ws in wb.worksheets:
        hr = None
        head = {}
        for r in range(1, min(ws.max_row, 10) + 1):
            names = {}
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if v is not None:
                    names[c] = str(v).strip()
            joined = ' '.join(names.values())
            if '表名称' in joined or ('名称' in joined and '类型' in joined):
                hr = r
                for c, n in names.items():
                    if '编号' in n:
                        head[c] = 'code'
                    elif '名称' in n:
                        head[c] = 'name'
                    elif n == '类型' or '类型' in n:
                        head[c] = 'kind'
                    elif '系统' in n:
                        head[c] = 'sys'
                    elif '区域' in n:
                        head[c] = 'zone'
                    elif '位置' in n:
                        head[c] = 'loc'
                    elif '柜' in n or '面板' in n:
                        head[c] = 'panel'
                    elif '用途' in n:
                        head[c] = 'purpose'
                    elif '倍率' in n:
                        head[c] = 'mult'
                    elif '层级' in n:
                        head[c] = 'level'
                    elif '主备' in n:
                        head[c] = 'use_status'
                    elif '折算' in n or '系数' in n:
                        head[c] = 'factor'
                    elif '总表' in n:
                        head[c] = 'is_total'
                    elif '水源' in n:
                        head[c] = 'source'
                    elif '国网' in n or '关口' in n:
                        head[c] = 'is_grid'
                    elif '启用' in n or '状态' in n:
                        head[c] = 'enabled'
                break
        if hr is None:
            continue
        for r in range(hr + 1, ws.max_row + 1):
            v = {f: ws.cell(r, c).value for c, f in head.items()}
            def _s(k):
                x = v.get(k)
                return '' if x is None else str(x).strip()
            name = _s('name')
            if not name:
                continue
            kind = 'water' if '水' in _s('kind') and '电' not in _s('kind') else 'elec'
            src = 'reclaimed' if '中水' in _s('source') else 'tap'
            grid = 1 if _s('is_grid') in ('是', '1', 'Y', 'y') else 0
            try:
                mult = float(v.get('mult') or 1)
            except (TypeError, ValueError):
                mult = 1
            # 计量层级：优先读「计量层级」列，兼容旧的「是否总表」列
            lv = _s('level')
            purpose = _s('purpose')
            if lv:
                is_total = 1 if '一级' in lv or lv in ('总表',) else 0
                is_sub_total = 1 if '二级' in lv else 0
                grid = 1 if ('关口' in lv or '国网' in lv) else 0
            else:
                is_total = 1 if _s('is_total') in ('是', '1') else 0
                is_sub_total = 1 if ('二级' in _s('is_total') or '区域总表' in _s('is_total')) else 0
            # 主备：优先读列，否则按名称/用途自动判定
            us = _s('use_status')
            if us in ('备用', 'backup'):
                use_status = 'backup'
            elif us in ('主用', 'main'):
                use_status = 'main'
            else:
                use_status = 'backup' if ('备用' in name or '备用' in purpose) else 'main'
            try:
                factor = float(v.get('factor') or 1) or 1
            except (TypeError, ValueError):
                factor = 1
            d = dict(code=_s('code'), name=name, kind=kind, sys=_s('sys') or '其他',
                     zone=_s('zone'), loc=_s('loc'), panel=_s('panel'), purpose=purpose,
                     mult=mult, is_total=is_total, is_sub_total=is_sub_total,
                     use_status=use_status, factor=factor,
                     enabled=0 if _s('enabled') in ('停用', '否', '0') else 1,
                     source=src, is_grid=grid)
            key = d['code'].strip() or ('N:' + d['name'])
            meta = existing.get(key)
            if meta:
                if update_existing:
                    meta.update(d)
                    meter_save(meta)
                    updated += 1
            else:
                meter_save(d)
                existing[key] = dict(id=0)
                added += 1
    return added, updated, skipped


# ---------- 抄表明细 导入/导出 ----------
def export_readings():
    try:
        from server.xlsx_io import write_xlsx
    except ImportError:
        from xlsx_io import write_xlsx
    cx = conn()
    data = [['抄表读数明细'], [], ['表编号', '表名称', '类型', '日期', '读数']]
    q = ('SELECT r.meter_id, r.rdate, r.val FROM reading r ORDER BY r.meter_id, r.rdate')
    mb = {m['id']: m for m in meters(None)}
    for row in cx.execute(q):
        m = mb.get(row['meter_id'])
        if not m:
            continue
        data.append([m['code'], m['name'], '电' if m['kind'] == 'elec' else '水',
                     row['rdate'], row['val']])
    cx.close()
    return write_xlsx(data, '读数', [12, 26, 6, 12, 14])


def import_readings(data):
    """导入抄表读数：列 = 表编号/表名称/类型/日期/读数（编号或名称匹配已有表）"""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    existing = {}
    for m in meters(None):
        if m['code']:
            existing[m['code'].strip()] = m
        existing['N:' + m['name']] = m
    reads = missed = 0
    buf = []
    for ws in wb.worksheets:
        hr = None
        head = {}
        for r in range(1, min(ws.max_row, 10) + 1):
            names = {c: str(ws.cell(r, c).value).strip()
                     for c in range(1, ws.max_column + 1) if ws.cell(r, c).value is not None}
            joined = ' '.join(names.values())
            if '读数' in joined and ('日期' in joined or '时间' in joined):
                hr = r
                for c, n in names.items():
                    if '编号' in n:
                        head[c] = 'code'
                    elif '名称' in n:
                        head[c] = 'name'
                    elif '日期' in n or '时间' in n:
                        head[c] = 'date'
                    elif '读数' in n or '示值' in n:
                        head[c] = 'val'
                break
        if hr is None:
            continue
        for r in range(hr + 1, ws.max_row + 1):
            code_c = next((c for c, f in head.items() if f == 'code'), None)
            code = '' if code_c is None or ws.cell(r, code_c).value is None else str(ws.cell(r, code_c).value).strip()
            name_c = next((c for c, f in head.items() if f == 'name'), None)
            name = '' if name_c is None or ws.cell(r, name_c).value is None else str(ws.cell(r, name_c).value).strip()
            date_c = next((c for c, f in head.items() if f == 'date'), None)
            val_c = next((c for c, f in head.items() if f == 'val'), None)
            if date_c is None or val_c is None:
                continue
            dstr = parse_date(ws.cell(r, date_c).value)
            val = to_num(ws.cell(r, val_c).value)
            if not dstr or val is None:
                missed += 1
                continue
            m = existing.get(code) or (existing.get('N:' + name) if name else None)
            if not m:
                missed += 1
                continue
            buf.append((m['id'], dstr, val))
            reads += 1
    if buf:
        cx = conn()
        cx.execute('PRAGMA busy_timeout=20000')
        cx.executemany(
            'INSERT INTO reading(meter_id,rdate,val,note,created) VALUES(?,?,?,?,?) '
            'ON CONFLICT(meter_id,rdate) DO UPDATE SET val=excluded.val',
            [(a, b, c, '', now()) for a, b, c in buf])
        cx.commit(); cx.close()
    return reads, missed


# ---------- 通用模板 ----------
TEMPLATE_HEAD = ['表编号', '表名称', '类型', '系统分类', '区域', '安装位置',
                 '2026-01-31', '2026-02-28', '2026-03-31']


def build_fill_template(month=None, days=None):
    """
    填空式模板（横向日期）：每块表一行，列=该月每个抄表日。
    两种形态：
      month=None → 汇总模式：每块表一行，列出最近 6 个账期列（按月），填每月月底读数
      month='2026-11' → 单月模式：列出该月每一天的列，日抄/周抄都填
    """
    try:
        from server.xlsx_io import write_xlsx
    except ImportError:
        from xlsx_io import write_xlsx
    cx = conn()
    ms = meters(None)
    # 读数索引 {(mid, 'YYYY-MM'): val, (mid, 'YYYY-MM-DD'): val}
    bym = {}
    byd = {}
    last = {}
    for row in cx.execute('SELECT meter_id, rdate, val FROM reading ORDER BY meter_id, rdate'):
        bym[(row['meter_id'], row['rdate'][:7])] = row['val']   # 该账期最后一次读数
        byd[(row['meter_id'], row['rdate'])] = row['val']
        last[row['meter_id']] = (row['rdate'], row['val'])
    cx.close()

    if month:
        import calendar as _cal
        y, mo = int(month[:4]), int(month[5:7])
        ndays = days or _cal.monthrange(y, mo)[1]
        cols = ['%s-%02d' % (month, d) for d in range(1, ndays + 1)]
        title = '抄表填空模板（%s 月）— 每块表一行，日期横向，只填要抄的格子' % month
    else:
        months_all = sorted({k[1] for k in bym.keys()})
        # 账期列的列名取该账期在库里的实际抄表日（多数是月底，也可能是月中），
        # 这样导入时按 (表,日期) 能精确命中已有读数，不会多插一条月底行
        day_of = {}
        for (mid_, mk), _v in bym.items():
            if mk not in day_of:
                day_of[mk] = mk + '-01'
        for (mid_, rd_) in byd.keys():
            mk = rd_[:7]
            if mk in day_of and rd_ > day_of[mk]:
                day_of[mk] = rd_
        months_all = months_all[-6:]
        cols = [day_of.get(mk, mk + '-01') for mk in months_all]
        title = '抄表填空模板（按月汇总）— 每块表一行 × 近 %d 个账期，填该账期最后一次读数' % len(cols)

    head = ['表ID(勿改)', '表编号', '表名称', '类型', '区域', '上次读数(参考)'] + cols
    data = [[title], [], head]
    for m in ms:
        if not m.get('enabled') or m.get('use_status') == 'backup':
            continue
        row = ['ID-%04d' % m['id'], m['code'] or '', m['name'],
               '电' if m['kind'] == 'elec' else '水', m['zone'] or '',
               last.get(m['id'], ('', None))[1]]
        for c in cols:
            # 单月模式按具体日期取；汇总模式列名是日期，取值用其所在账期
            v = byd.get((m['id'], c)) if month else bym.get((m['id'], c[:7]))
            row.append(v)
        data.append(row)
    data.append([])
    data.append(['填写说明：'])
    if month:
        data.append(['1. 只在表格里填读数；不抄的日期留空；一行可填多个日期（周抄/日抄都行）'])
    else:
        data.append(['1. 每列一个账期，填该账期最后一次抄表读数（月中抄多次时填最后一次即可）'])
    data.append(['2. A 列表ID用于精确匹配，请勿改动；已有读数会显示在格子里，方便对照'])
    data.append(['3. 导入时会预检并让你逐行确认，不会直接覆盖数据'])
    data.append(['4. 新增的表请先到「表计台账」建档，再重新导出本模板'])
    widths = [12, 12, 30, 6, 10, 14] + [12] * len(cols)
    return write_xlsx(data, '填空模板', widths)


FILL_PAT = re.compile(r'^ID-(\d+)$')


def _cell_text(ws, r, c):
    """取单元格文本；日期型单元格格式化为 YYYY-MM-DD（月份列按月底处理）"""
    v = ws.cell(r, c).value
    if v is None:
        return ''
    if hasattr(v, 'strftime'):
        # Excel 日期：月份列常为当月 1 号，这里保持原样，交给 parse_date 处理
        return v.strftime('%Y-%m-%d')
    return str(v).strip()


def _header_texts(ws, hr):
    """表头行文字映射 {列号: 文本}，供横向模板取账期列名"""
    out = {}
    for c in range(1, ws.max_column + 1):
        t = _cell_text(ws, hr, c)
        if t:
            out[c] = t
    return out


def _extract_rows(ws, head_map, hr, wide, header_row=None):
    """
    从工作表抽取读数行。
    wide=True（横向模板）：每行=一块表，日期列有多个，逐列生成 (mid, date, val)
                           日期取该列的**表头文字**（如 2026-10 / 2026-10-31）
    wide=False（纵向模板）：每行=一条读数
    """
    out = []
    mid_c = next((c for c, f in head_map.items() if f == 'mid'), None)
    code_c = next((c for c, f in head_map.items() if f == 'code'), None)
    name_c = next((c for c, f in head_map.items() if f == 'name'), None)
    if wide:
        htext = header_row if header_row is not None else _header_texts(ws, hr)
        date_cols = [c for c, f in head_map.items() if f == 'date']
        for r in range(hr + 1, ws.max_row + 1):
            midv = ws.cell(r, mid_c).value if mid_c else None
            code = '' if code_c is None or ws.cell(r, code_c).value is None else str(ws.cell(r, code_c).value).strip()
            name = '' if name_c is None or ws.cell(r, name_c).value is None else str(ws.cell(r, name_c).value).strip()
            if not (midv or code or name):
                continue
            for c in sorted(date_cols):
                v = to_num(ws.cell(r, c).value)
                if v is None:
                    continue
                dstr = htext.get(c) or ''       # 列名即账期/日期
                ds = ''
                t = str(dstr).strip()
                # 月粒度（2026-11 / 2026.11）优先判 → 该月最后一天（月底读数惯例）
                m2 = re.match(r'^(\d{4})[-/.](\d{1,2})$', t)
                if m2:
                    import calendar as _cal
                    y2, mo2 = int(m2.group(1)), int(m2.group(2))
                    if 1 <= mo2 <= 12:
                        ds = '%04d-%02d-%02d' % (y2, mo2, _cal.monthrange(y2, mo2)[1])
                if not ds:
                    ds = parse_date(dstr)
                if ds:
                    out.append((r, midv, code, name, ds, v))
    else:
        date_c = next((c for c, f in head_map.items() if f == 'date'), None)
        val_c = next((c for c, f in head_map.items() if f == 'val'), None)
        if date_c is None or val_c is None:
            return out
        for r in range(hr + 1, ws.max_row + 1):
            midv = ws.cell(r, mid_c).value if mid_c else None
            code = '' if code_c is None or ws.cell(r, code_c).value is None else str(ws.cell(r, code_c).value).strip()
            name = '' if name_c is None or ws.cell(r, name_c).value is None else str(ws.cell(r, name_c).value).strip()
            if not (midv or code or name):
                continue
            dstr = parse_date(ws.cell(r, date_c).value)
            v = to_num(ws.cell(r, val_c).value)
            if dstr and v is not None:
                out.append((r, midv, code, name, dstr, v))
    return out


def _detect_head(ws):
    """识别表头，返回 (hr, head_map, wide)；wide=True 为横向日期模板"""
    for r in range(1, min(ws.max_row, 10) + 1):
        names = {c: str(ws.cell(r, c).value).strip()
                 for c in range(1, ws.max_column + 1) if ws.cell(r, c).value is not None}
        joined = ' '.join(names.values())
        if '读数' not in joined and '日期' not in joined:
            continue
        head = {}
        date_cnt = 0
        for c, n in names.items():
            if n.startswith('ID') or '表ID' in n:
                head[c] = 'mid'
            elif '编号' in n:
                head[c] = 'code'
            elif '名称' in n:
                head[c] = 'name'
            elif n == '读数' or '读数(度)' in n or '读数（度）' in n or '本次读数' in n:
                head[c] = 'val'
            elif ('本次' in n or '抄表日' in n or '日期' in n or '时间' in n
                  or re.match(r'^\d{4}[-/.]\d{1,2}([-/.]\d{1,2})?$', n)):
                head[c] = 'date'
                date_cnt += 1
        # 有表ID + 多个日期列 → 横向
        wide = 'mid' in head.values() and date_cnt > 1
        has_subject = ('mid' in head.values() or 'code' in head.values()
                       or 'name' in head.values())
        if wide:
            if has_subject:
                return r, head, wide
            continue
        # 纵向：必须同时有「日期」和「读数」两列，避免把标题行当表头
        if has_subject and 'date' in head.values() and 'val' in head.values():
            return r, head, wide
    return None, {}, False


def precheck_readings(data):
    """
    读数导入预检：不写库。支持横向（一块表一行×日期列）与纵向两种格式。
    返回 {rows: [...], stats: {...}}
    每行额外带 exists（库里该日期是否已有读数）供前端标注「将覆盖/新增」。
    """
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    existing = {}
    for m in meters(None):
        existing['ID-%04d' % m['id']] = m
        if m['code']:
            existing.setdefault(m['code'].strip(), m)
        existing.setdefault('N:' + m['name'], m)
    cx = conn()
    last = {}
    hist = {}
    cur_vals = {}
    for row in cx.execute('SELECT meter_id, rdate, val FROM reading ORDER BY meter_id, rdate'):
        last[row['meter_id']] = (row['rdate'], row['val'])
        hist.setdefault(row['meter_id'], []).append((row['rdate'], row['val']))
        cur_vals[(row['meter_id'], row['rdate'])] = row['val']
    cx.close()
    prev2 = {}
    for mid, h in hist.items():
        if len(h) >= 2:
            prev2[mid] = h[-2]

    rows_out = []
    n_ok = n_warn = n_err = n_overwrite = n_same = 0
    rid = 0
    for ws in wb.worksheets:
        hr, head, wide = _detect_head(ws)
        if hr is None:
            continue
        for (row, midv, code, name, dstr, val) in _extract_rows(
                ws, head, hr, wide, header_row=_header_texts(ws, hr)):
            key_m = 'ID-%04d' % int(midv[3:]) if (midv and FILL_PAT.match(str(midv).strip())) else (str(midv).strip() if midv else '')
            m = (existing.get(key_m) if key_m else None) or existing.get(code) or (
                existing.get('N:' + name) if name else None)
            rid += 1
            item = dict(rid=rid, row=row, mid=(m or {}).get('id'),
                        name=(m or {}).get('name') or (name or code or str(midv or '')),
                        code=(m or {}).get('code') or code,
                        date=dstr, val=val,
                        last=None, last_date=None, exists=None,
                        ok=False, warn=[], level='', msg='')
            if not m:
                item['level'] = 'err'
                item['msg'] = '找不到对应表计（表ID/编号/名称均未匹配），请先在台账建档'
                n_err += 1
            elif val is None:
                item['level'] = 'skip'
                item['msg'] = '读数为空'
            else:
                ek = (m['id'], dstr)
                if ek in cur_vals:
                    item['exists'] = cur_vals[ek]
                    if abs((cur_vals[ek] or 0) - (val or 0)) < 1e-9:
                        # 与库中值完全一致 → 无需导入
                        item['level'] = 'skip'
                        item['msg'] = '与现有读数相同，无需导入'
                        n_same += 1
                        rows_out.append(item)
                        continue
                    n_overwrite += 1
                la = last.get(m['id'])
                item['last'] = la[1] if la else None
                item['last_date'] = la[0] if la else None
                p2 = prev2.get(m['id'])
                if la and la[1] is not None and val < la[1]:
                    drop = (la[1] - val) / la[1] * 100 if la[1] else 0
                    if drop > 80 and (p2 is None or val < p2[1]):
                        item['level'] = 'warn'
                        item['msg'] = '读数从 %s 降到 %s（降幅 %.0f%%），疑似换表/表底清零' % (la[1], val, drop)
                        item['warn'].append('换表?')
                    else:
                        item['level'] = 'warn'
                        item['msg'] = '读数小于上次读数 %s（倒走），请确认是否抄错' % la[1]
                        item['warn'].append('倒走')
                    n_warn += 1
                elif la and la[1] is not None:
                    inc = val - la[1]
                    try:
                        from datetime import date as _dd
                        d0 = _dd(*map(int, la[0].split('-')))
                        d1 = _dd(*map(int, dstr.split('-')))
                        days = max((d1 - d0).days, 1)
                    except Exception:
                        days = 30
                    daily = inc / days
                    ref = la[1] / max(days, 1)
                    if inc > 0 and la[1] > 10 and ref and daily > ref * 8:
                        item['level'] = 'warn'
                        item['msg'] = '较上次 +%s，日均暴涨约 %.0f 倍，请确认' % (round(inc, 1), daily / ref)
                        item['warn'].append('暴涨')
                        n_warn += 1
                    else:
                        item['level'] = 'ok'
                        item['msg'] = ''
                        n_ok += 1
                else:
                    item['level'] = 'ok'
                    item['msg'] = '该表首次读数'
                    n_ok += 1
            if item['exists'] is not None and item['level'] == 'ok':
                item['msg'] = '该日期已有读数 %s，导入将覆盖（可取消勾选保留原值）' % item['exists']
            item['ok'] = item['level'] == 'ok'
            rows_out.append(item)
    return dict(rows=rows_out,
                stats=dict(total=len(rows_out), ok=n_ok, warn=n_warn, err=n_err,
                           overwrite=n_overwrite, same=n_same))


def import_readings_checked(data, skip_rids=None, overwrite=True):
    """
    按预检结果导入。
    skip_rids：不导入的行号集合
    overwrite=False：库中该 (表,日期) 已有读数的行跳过不覆盖
    返回 (writes, skipped, kept(保留原值), missed, same(无需变更))
    """
    skip = set(skip_rids or [])
    pc = precheck_readings(data)
    buf = []
    skipped = kept = missed = same = 0
    for it in pc['rows']:
        if it['rid'] in skip:
            skipped += 1
            continue
        if it['level'] == 'skip':
            # 与库中完全一致（或读数为空）→ 不写库，计入「无需变更」
            same += 1
            continue
        if it['level'] == 'err':
            missed += 1
            continue
        if it['exists'] is not None and not overwrite:
            kept += 1
            continue
        if it['mid'] and it['date'] and it['val'] is not None:
            buf.append((it['mid'], it['date'], it['val']))
        else:
            missed += 1
    if buf:
        cx = conn()
        cx.execute('PRAGMA busy_timeout=20000')
        cx.executemany(
            'INSERT INTO reading(meter_id,rdate,val,note,created) VALUES(?,?,?,?,?) '
            'ON CONFLICT(meter_id,rdate) ' + ('DO UPDATE SET val=excluded.val' if overwrite
                                              else 'DO NOTHING'),
            [(a, b, c, '', now()) for a, b, c in buf])
        cx.commit(); cx.close()
    return len(buf), skipped, kept, missed, same


def build_template(rows=None):
    try:
        from server.xlsx_io import write_xlsx
    except ImportError:
        from xlsx_io import write_xlsx
    data = [['能耗抄表导入模板'], [], TEMPLATE_HEAD]
    if rows:
        data.extend(rows)
    else:
        data.append(['E-D101', '示例-商铺照明01AL1', '电', '商铺照明', 'A区', 'A区负一层', 1042.7, 1259.8, 1270])
        data.append(['W-001', '示例-A区女卫', '水', '卫生间', 'A区', 'A区女卫', 68, 100, 115])
    data.append([])
    data.append(['填写说明：'])
    data.append(['1. 类型填“电”或“水”；表编号相同即视为同一块表，重复导入会自动更新'])
    data.append(['2. 日期列可自行增删，格式 2026-01-31 或 2026.1.31 均可'])
    data.append(['3. 系统分类可填：' + '、'.join(['照明', '商铺照明', '景观照明', '消防', '电梯扶梯', '空调用电', '排水泵', '动力排风', '安防', '临电转供', '卫生间', '绿化灌溉', '绿化雾森', '水景', '消防补水', '中水补水', '设备间', '其他'])])
    data.append(['4. 每月只需填当月读数，上月读数由系统保存；漏填的格子留空即可'])
    return write_xlsx(data, '导入模板', [16, 26, 8, 12, 10, 22, 12, 12, 12])
