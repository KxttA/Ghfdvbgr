# -*- coding: utf-8 -*-
"""能耗报表导出：多 sheet、带图表、带同比环比、含年度总表与分组趋势"""
import io, json, datetime
try:
    from server.xlsx_io import build_workbook
    from server import db as DB
except ImportError:
    from xlsx_io import build_workbook
    import db as DB


def _all_dicts():
    c = DB.conn()
    r = [dict(x) for x in c.execute('SELECT * FROM dict ORDER BY kind,sort,id')]
    c.close()
    return r


def _n(v, d=1):
    return '—' if v is None else ('%.' + str(d) + 'f') % v


def _p(v):
    return '—' if v is None else '%+.1f%%' % v


def _coll(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build(a, kind, mk, year=None):
    kw = '电' if kind == 'elec' else '水'
    unit = 'kWh' if kind == 'elec' else '吨'
    price = a.price_of(kind, mk)
    y, mo = mk[:4], int(mk[5:7])
    year = int(year or y)
    sheets = []

    # ---------- 1 概览 ----------
    cmp = a.compare(kind, mk)
    cov = a.coverage(kind, mk)
    ov = [
        ['能耗分析报表（%s）— %s 年 %d 月' % (kw, y, mo)],
        [],
        ['指标', '本月', '上月', '环比', '去年同期', '同比'],
        ['用量（%s）' % unit, _n(cmp['cur']), _n(cmp['prev']) if cmp['prev'] is not None else '—',
         _p(cmp['mom']), _n(cmp['last_year']), _p(cmp['yoy'])],
        ['日均用量（%s/天）' % unit, _n(cmp['daily_cur'], 2),
         _n(cmp['daily_prev'], 2) if cmp['daily_prev'] else '—', _p(cmp['mom_daily']), '—', '—'],
        ['费用（元）', _n(cmp['cost']), '—', '—', '—', '—'],
        ['本月单价（元/%s）' % ('度' if kind == 'elec' else '吨'), '%.4f' % price, '', '', '', ''],
        ['抄表覆盖（块表）', '%d / %d（%.1f%%）' % (cov['covered'], cov['total'], cov['pct']), '', '', '', ''],
    ]
    _st = a.total(kind, [mk])
    _sb = a.total(kind, [mk], basis='sub')
    _tv = _st.get(mk, {}).get('usage', 0)
    _sv = _sb.get(mk, {}).get('usage', 0)
    ov.append(['分表合计（%s）' % unit, _n(_sv), '', '', '', ''])
    ov.append(['总表−分表 差额', _n(_tv - _sv), '',
               ('%.1f%%' % ((_tv - _sv) / _sv * 100)) if _sv else '—', '', ''])
    ov += [
        [],
        ['口径说明：'],
        ['1. 总用量按总表统计：电=各配电室低压进线柜总表合计，水=自来水总表（中水单独核算）'],
        ['2. 总表/二级总表不参与分项合计，分表合计仅用于损耗与计量校核，避免重复计量'],
        ['3. 备用回路不计入合计；用量 = 抄见用量 × 折算系数（默认 1）'],
        ['4. 完整计算公式见「计算公式」工作表'],
    ]
    charts_ov = []
    st0 = a.structure(kind, mk, 'sys')
    if st0:
        for i in st0:
            ov.append([i['label'], i['usage'], '%.1f%%' % i['share'], '', '', ''])
        n = len(st0)
        charts_ov.append(dict(kind='bar', title='%s 年 %d 月 分系统用量（%s）' % (y, mo, unit),
                              cat='A15:A%d' % (14 + n), cat_vals=[i['label'] for i in st0],
                              sers=[('用量', 'B15:B%d' % (14 + n), [i['usage'] for i in st0])],
                              anchor='H3', size=(12, 18)))
    sheets.append(dict(title='分析概览', rows=ov, widths=[24, 18, 14, 12, 14, 12], charts=charts_ov))

    # ---------- 2 年度每月总表（可视化） ----------
    ys = a.year_summary(kind, year)
    rows = [['%s 年度每月用量总表（%s）' % (year, kw)], [],
            ['月份', '用量（%s）' % unit, '日均（%s/天）' % unit, '单价（元）', '费用（元）', '环比', '较月均']]
    avg = ys['avg_month']
    prevu = None
    for r in ys['rows']:
        mom = None
        if r['usage'] is not None and prevu:
            mom = (r['usage'] - prevu) / prevu * 100
        vs = ((r['usage'] - avg) / avg * 100) if (r['usage'] is not None and avg) else None
        rows.append(['%d月' % int(r['month'][5:]),
                     _n(r['usage']),
                     _n(r['daily'], 2) if r['usage'] is not None else '—',
                     '%.4f' % r['price'],
                     _n(r['cost'], 0) if r['cost'] is not None else '—',
                     _p(mom), _p(vs)])
        if r['usage']:
            prevu = r['usage']
    rows.append([])
    rows.append(['全年合计', _n(ys['total_usage']), '—', '—', _n(ys['total_cost'], 0), '—', '—'])
    rows.append(['月均用量', _n(ys['avg_month']), '—', '—', '—', '—', '—'])
    rows.append(['有数据月份', '%d 个月' % ys['months_count'], '—', '—', '—', '—', '—'])
    charts_y = []
    if ys['months_count']:
        n = len(ys['rows'])
        charts_y.append(dict(kind='col', title='%s 年每月用量（%s）' % (year, unit),
                             cat='A4:A%d' % (3 + n),
                             cat_vals=['%d月' % i for i in range(1, 13)],
                             sers=[('用量', 'B4:B%d' % (3 + n),
                                    [r['usage'] if r['usage'] is not None else 0 for r in ys['rows']])],
                             anchor='I4', size=(14, 20)))
    sheets.append(dict(title='年度每月总表', rows=rows, widths=[12, 16, 16, 12, 14, 12, 12], charts=charts_y))

    # ---------- 3 月度趋势（全部月份） ----------
    tr = []
    tt = a.total(kind)
    for m in a.months:
        v = tt.get(m, {})
        tr.append(dict(month=m, usage=v.get('usage', 0), cost=v.get('cost', 0),
                       daily=round(v.get('usage', 0) / DB.days_in_month(m), 2)))
    rows = [['月度用量趋势（%s）' % kw], [],
            ['月份', '用量（%s）' % unit, '日均（%s/天）' % unit, '费用（元）', '环比']]
    prevu = None
    for t in tr:
        mom = None if prevu in (None, 0) else (t['usage'] - prevu) / prevu * 100
        rows.append([t['month'], round(t['usage'], 1), round(t['daily'], 2),
                     round(t['cost'], 0), _p(mom)])
        prevu = t['usage'] or None
    charts = []
    if len(tr) >= 2:
        charts.append(dict(kind='col', title='月度用量（%s）' % unit,
                           cat='A4:A%d' % (3 + len(tr)),
                           cat_vals=[t['month'] for t in tr],
                           sers=[('用量', 'B4:B%d' % (3 + len(tr)), [round(t['usage'], 1) for t in tr])],
                           anchor='G4', size=(12, 17)))
        charts.append(dict(kind='line', title='月度日均用量（%s/天）' % unit,
                           cat='A4:A%d' % (3 + len(tr)),
                           cat_vals=[t['month'] for t in tr],
                           sers=[('日均', 'C4:C%d' % (3 + len(tr)), [round(t['daily'], 2) for t in tr])],
                           anchor='G23', size=(12, 17)))
    sheets.append(dict(title='月度趋势', rows=rows, widths=[14, 16, 16, 14, 12], charts=charts))

    # ---------- 4 分组趋势（A+D+北通道 / E+F 等，自由设置） ----------
    groups = [g for g in DB.groups() if g.get('kind') in ('both', kind)]
    gs = a.group_series(kind, a.months, groups)
    if gs:
        rows = [['分区/分单元用量对比（%s）' % kw], [],
                ['月份'] + [g['name'] for g in gs] + ['合计']]
        for i, m in enumerate(a.months):
            vals = [g['months'][i]['usage'] for g in gs]
            rows.append([m] + [round(v, 1) for v in vals] + [round(sum(vals), 1)])
        rows.append(['合计'] + [g['total'] for g in gs] + [round(sum(g['total'] for g in gs), 1)])
        rows.append([])
        rows.append(['分组口径：'])
        for g in gs:
            rows.append([g['name'],
                         '区域：' + ('、'.join(g['zones']) or '全部') +
                         ('；系统：' + '、'.join(g['sys']) if g['sys'] else '')])
        charts = []
        if len(a.months) >= 2:
            n = len(a.months)
            charts.append(dict(kind='col', title='各分组月度用量对比（%s）' % unit,
                               cat='A4:A%d' % (3 + n), cat_vals=a.months,
                               sers=[(g['name'], '%s4:%s%d' % (_coll(2 + j), _coll(2 + j), 3 + n),
                                      [round(g['months'][i]['usage'], 1) for i in range(n)])
                                     for j, g in enumerate(gs)],
                               anchor='H4', size=(14, 20)))
        sheets.append(dict(title='分组趋势', rows=rows,
                           widths=[14] + [16] * (len(gs) + 1), charts=charts))

    # ---------- 5 分系统结构 ----------
    st = a.structure(kind, mk, 'sys')
    rows = [['分系统能耗结构 — %s 年 %d 月' % (y, mo)], [],
            ['系统分类', '用量（%s）' % unit, '占比', '上月用量', '环比', '同比', '费用（元）', '费用变化（元）']]
    for i in st:
        rows.append([i['label'], i['usage'], '%.1f%%' % i['share'],
                     _n(i['prev']) if i['prev'] is not None else '—', _p(i['mom']), _p(i['yoy']),
                     i.get('cost', 0), i.get('cost_delta', 0)])
    rows.append(['合计', sum(i['usage'] for i in st), '100.0%', '', '', '',
                 round(sum(i['usage'] for i in st) * price, 0), ''])
    charts = []
    if st:
        n = len(st)
        charts.append(dict(kind='bar', title='各系统用量（%s）' % unit,
                           cat='A4:A%d' % (3 + n), cat_vals=[i['label'] for i in st],
                           sers=[('本月', 'B4:B%d' % (3 + n), [i['usage'] for i in st])],
                           anchor='J4', size=(12, 18)))
        charts.append(dict(kind='col', title='各系统环比（%）',
                           cat='A4:A%d' % (3 + n), cat_vals=[i['label'] for i in st],
                           sers=[('环比', 'E4:E%d' % (3 + n), [i['mom'] or 0 for i in st])],
                           anchor='J24', size=(12, 15)))
    sheets.append(dict(title='分系统结构', rows=rows, widths=[18, 16, 10, 14, 12, 12, 16, 16], charts=charts))

    # ---------- 6 分区结构 ----------
    zn = a.structure(kind, mk, 'zone')
    rows = [['分区域能耗结构 — %s 年 %d 月' % (y, mo)], [],
            ['区域', '用量（%s）' % unit, '占比', '上月用量', '环比', '同比', '费用（元）']]
    for i in zn:
        rows.append([i['label'], i['usage'], '%.1f%%' % i['share'],
                     _n(i['prev']) if i['prev'] is not None else '—', _p(i['mom']), _p(i['yoy']),
                     i.get('cost', 0)])
    rows.append(['合计', sum(i['usage'] for i in zn), '100.0%', '', '', '',
                 round(sum(i['usage'] for i in zn) * price, 0)])
    charts = []
    if zn:
        n = len(zn)
        charts.append(dict(kind='bar', title='各区域用量（%s）' % unit,
                           cat='A4:A%d' % (3 + n), cat_vals=[i['label'] for i in zn],
                           sers=[('本月', 'B4:B%d' % (3 + n), [i['usage'] for i in zn])],
                           anchor='I4', size=(12, 18)))
    sheets.append(dict(title='分区域结构', rows=rows, widths=[14, 16, 10, 14, 12, 12, 16], charts=charts))

    # ---------- 7 表计明细 ----------
    rows = [['表计用量明细 — %s 年 %d 月' % (y, mo)], [],
            ['表计名称', '系统', '区域', '用量（%s）' % unit, '上月用量', '环比',
             '费用（元）', '是否总表']]
    for t in a.top_meters(kind, mk, 999):
        rows.append([t['name'], t['sys'], t['zone'], t['usage'],
                     _n(t['prev']) if t['prev'] is not None else '—', _p(t['mom']),
                     t['cost'], '是' if t['is_total'] else ''])
    sheets.append(dict(title='表计明细', rows=rows, widths=[30, 12, 10, 14, 14, 10, 14, 10]))

    # ---------- 8 异常 ----------
    an = a.anomalies(kind, mk)
    rows = [['异常与风险提示 — %s' % mk], [],
            ['等级', '类型', '表计', '情况', '影响']]
    if an:
        for x in an:
            rows.append([x['level'], x['type'], x['meter'], x['detail'], x['impact']])
    else:
        rows.append(['—', '无异常', '—', '本月未发现异常', '—'])
    sheets.append(dict(title='异常提示', rows=rows, widths=[8, 14, 26, 46, 36]))

    # ---------- 9 节能建议 ----------
    sv = a.savings(kind, mk)
    rows = [['节能建议与测算 — %s 年 %d 月' % (y, mo)], [],
            ['序号', '建议措施', '对应系统', '判断依据', '具体做法', '预计节降量（%s）' % unit,
             '预计节降金额（元）', '节降比例']]
    for i, s in enumerate(sv, 1):
        rows.append([i, s['title'], s['sys'], s['basis'], s['advice'],
                     s['save_usage'], s['save_cost'], '%s%%' % s['ratio']])
    sheets.append(dict(title='节能建议', rows=rows, widths=[6, 26, 14, 34, 60, 16, 18, 10]))

    # ---------- 10 抄表质量 ----------
    q = a.quality(kind)
    rows = [['抄表质量 — %s' % kw], [],
            ['表计', '系统', '区域', '抄表次数', '平均间隔（天）', '最大间隔（天）', '状态']]
    for x in q:
        rows.append([x['name'], x.get('sys', ''), x.get('zone', ''), x['cnt'],
                     x['avg_gap'], x['max_gap'], x['status']])
    sheets.append(dict(title='抄表质量', rows=rows, widths=[30, 12, 10, 12, 16, 16, 12]))

    # ---------- 11 中水核算（仅水） ----------
    if kind == 'water':
        rc = a.reclaimed()
        rows = [['中水核算与分摊 — 截至 %s' % (a.months[-1] if a.months else '—')], [],
                ['指标', '数量（吨）']]
        rows.append(['中水补水量（接自来水）', rc['supply']])
        rows.append(['中水出水合计（各区绿化灌溉等）', rc['reclaimed']])
        rows.append(['中水损耗（补水-出水）', rc['loss']])
        rows.append(['损耗率', ('%.1f%%' % rc['loss_pct']) if rc['loss_pct'] is not None else '—'])
        rows.append([])
        rows.append(['区域', '中水用量（吨）', '占比', '分摊补水量（吨）',
                     '分摊算法：区中水用量 ÷ 中水总量 × 中水补水量'])
        for z in rc['zones']:
            rows.append([z['zone'], z['usage'], '%.1f%%' % z['share'], z['alloc'], ''])
        rows.append([])
        rows.append(['月份', '补水量（吨）', '中水用量（吨）', '损耗（吨）'])
        for r2 in rc['by_month']:
            rows.append([r2['month'], r2['supply'], r2['reclaimed'], r2['loss']])
        sheets.append(dict(title='中水核算', rows=rows, widths=[16, 20, 14, 20, 44]))

    # ---------- 12 损耗分析（三类总表比对 + 国网核对） ----------
    ls = a.loss(kind, a.months[-12:])
    rows = [['损耗与三类总表比对（%s）— 近 12 个月' % kw], [],
            ['口径：能耗分析以「②一级总表（低压配电柜进线柜 D101/D201）」为用量口径，最准确。'],
            ['①国网关口表仅用于核对自抄误差；③分表合计仅用于衡量未计量部分；三者不可相加。'],
            ['变损线损 = ① − ②（变压器/高压线路损耗）；未计量 = ② − ③；总损耗率 =（① − ③）÷ ①。'], []]
    if kind == 'elec':
        rows.append(['月份', '①国网结算（度）', '①环比', '①同比',
                     '②一级总表（度）', '②环比', '②同比',
                     '③分表合计（度）', '③环比', '③同比',
                     '①−② 变损线损（度）', '①−② 差率', '①−② 环比',
                     '②−③ 未计量（度）', '②−③ 差率', '②−③ 环比',
                     '总损耗（度）', '总损耗率', '损耗环比', '损耗同比'])
        for r2 in ls['rows']:
            rows.append([r2['month'],
                         r2['g'] or None, _p(r2.get('g_mom')), _p(r2.get('g_yoy')),
                         r2['l1'], _p(r2.get('l1_mom')), _p(r2.get('l1_yoy')),
                         r2['sub'], _p(r2.get('sub_mom')), _p(r2.get('sub_yoy')),
                         r2.get('d_g_l1'), _n(r2.get('d_g_l1_pct')) + '%' if r2.get('d_g_l1_pct') is not None else '—',
                         _p(r2.get('d_g_l1_mom')),
                         r2.get('d_l1_sub'), _n(r2.get('d_l1_sub_pct')) + '%' if r2.get('d_l1_sub_pct') is not None else '—',
                         _p(r2.get('d_l1_sub_mom')),
                         r2.get('d_g_sub'), _n(r2.get('d_g_sub_pct')) + '%' if r2.get('d_g_sub_pct') is not None else '—',
                         _p(r2.get('d_g_sub_mom')), _p(r2.get('d_g_sub_yoy'))])
    else:
        rows.append(['月份', '①自来水总表（吨）', '①环比', '①同比',
                     '②分表合计（吨）', '②环比', '②同比',
                     '①−② 损耗（吨）', '损耗率', '损耗环比', '损耗同比'])
        for r2 in ls['rows']:
            rows.append([r2['month'], r2['l1'], _p(r2.get('l1_mom')), _p(r2.get('l1_yoy')),
                         r2['sub'], _p(r2.get('sub_mom')), _p(r2.get('sub_yoy')),
                         r2.get('d_l1_sub'),
                         _n(r2.get('d_l1_sub_pct')) + '%' if r2.get('d_l1_sub_pct') is not None else '—',
                         _p(r2.get('d_l1_sub_mom')), _p(r2.get('d_l1_sub_yoy'))])
    rows += [[], ['国网抄表核对（自抄一级表 vs 国网）'], [],
             ['账期', '国网结算', '自抄一级表', '差值（国网−自抄）', '误差率', '差异环比', '差异同比', '国网抄表日', '备注']]
    has = False
    for c in ls.get('check', []):
        if c['grid_usage'] is None:
            continue
        has = True
        rows.append([c['month'], c['grid_usage'], c['mine'], c['diff'],
                     ('%.2f%%' % c['gap']) if c['gap'] is not None else '—',
                     _p(c.get('mom')), _p(c.get('yoy')), c['rdate'], c['note']])
    if not has:
        rows.append(['—', '未录入国网抄表', '', '', '', '', '', '', '在「损耗分析」页逐月录入'])

    # 各级总表明细（含环比/同比）
    if ls.get('detail'):
        rows += [[], ['各级总表逐月用量明细'], [],
                 ['表计（层级）'] + list(ls['months']) + ['合计', '最新环比', '最新同比']]
        for dt in ls['detail']:
            nm = ('【国网关口】' if dt['grid'] else '【一级总表】') + dt['name']
            rows.append([nm] + dt['data'] + [dt.get('total'),
                         _p((dt.get('mom') or [None])[-1]), _p((dt.get('yoy') or [None])[-1])])

    # 汇总：一级表 vs 分表
    try:
        stt = a.total(kind, a.months[-12:])
        sbt = a.total(kind, a.months[-12:], basis='sub')
        rows += [[], ['汇总：一级总表 vs 分表'], [], ['月份', '一级总表（分析口径）', '分表合计', '差额', '差额率']]
        for mk2 in a.months[-12:]:
            tv = stt.get(mk2, {}).get('usage', 0)
            sv = sbt.get(mk2, {}).get('usage', 0)
            rows.append([mk2, tv, sv, round(tv - sv, 1),
                         ('%.1f%%' % ((tv - sv) / sv * 100)) if sv else '—'])
    except Exception:
        pass
    sheets.append(dict(title='损耗分析', rows=rows,
                       widths=[20, 15, 10, 10, 15, 10, 10, 15, 10, 10, 15, 10, 10, 15, 10, 10, 13, 11, 10, 10]))

    # ---------- 13 计算公式 ----------
    sheets.append(dict(title='计算公式', rows=_formula_rows(kind, mk, a), 
                       widths=[6, 24, 58, 26]))

    return build_workbook(sheets)


def _formula_rows(kind, mk, a):
    unit = 'kWh（度）' if kind == 'elec' else '吨'
    st = a.total(kind, [mk]).get(mk, {})
    sb = a.total(kind, [mk], basis='sub').get(mk, {})
    tm = (kind == 'elec')
    rows = [
        ['计算公式与口径说明（%s）' % ('电' if tm else '水')],
        [],
        ['序号', '计算项', '公式', '示例（%s）' % mk],
        [1, '单表某段用量', '本次读数 − 上次读数（若设倍率则 × 倍率）', ''],
        [2, '折算系数', '实际用量 = 抄见用量 × 折算系数（默认 1）', ''],
        [3, '用量归月', '该段用量 × 该段在本月的天数 ÷ 该段总天数（按天均摊到各自然月）', ''],
        [4, '总用量（一级总表）',
         ('各配电室低压进线柜总表用量之和（D101 + D201 + …）' if tm
          else '自来水总表用量（中水不计入）'), '%s %s' % (_n(st.get('usage', 0)), unit)],
        [5, '分表合计',
         '所有末端分表用量之和（不含总表、二级总表、关口表、备用表%s）' % ('、中水表' if not tm else ''),
         '%s %s' % (_n(sb.get('usage', 0)), unit)],
        [6, '总表−分表 差额', '总用量 − 分表合计', '%s %s' % (_n(st.get('usage', 0) - sb.get('usage', 0)), unit)],
        [7, '总损耗率', '（总用量 − 分表合计）÷ 总用量 × 100%',
         ('%.1f%%' % ((st.get('usage', 0) - sb.get('usage', 0)) / st.get('usage') * 100))
         if st.get('usage') else '—'],
    ]
    if tm:
        rows += [
            [8, '变损线损（一次损耗）', '国网结算电量 − 自抄一级总表合计', ''],
            [9, '未计量（二次损耗）', '一级总表合计 − 分表合计', ''],
            [10, '国网误差', '国网结算电量 − 自抄总表用量'],
            [11, '国网误差率', '（国网结算 − 自抄总表）÷ 国网结算 × 100%', ''],
        ]
    else:
        rows += [
            [8, '中水补水量', '「中水补水」分类的表用量合计（该表接自来水）', ''],
            [9, '中水出水合计', '所有中水表（source=中水）用量之和', ''],
            [10, '某区应分摊自来水', '该区中水用量 ÷ 所有中水表用量和 × 中水补水量', ''],
            [11, '中水损耗', '中水补水量 − 中水出水合计', ''],
        ]
    rows += [
        [],
        ['', '环比', '（本月 − 上月）÷ 上月 × 100%'],
        ['', '日均环比', '（本月日均 − 上月日均）÷ 上月日均 × 100%；日均 = 用量 ÷ 当月天数'],
        ['', '同比', '（本月 − 去年同月）÷ 去年同月 × 100%；无去年数据留空'],
        ['', '费用', '用量 × 当月单价（单价可在「系统设置 → 按月电价/水价」逐月设置）'],
        ['', '占比', '该项目用量 ÷ 同期合计用量 × 100%'],
        [],
        ['层级定义：'],
        ['', '项目关口总表', '国网结算表（高压侧），用于与自抄总表比对误差'],
        ['', '一级总表', '配电室低压进线柜总表（电）/ 自来水总表（水）——总用量口径'],
        ['', '二级总表', '区域/楼栋总表（如 C区、BDEF区总水表），隶属一级总表，仅兜底不叠加'],
        ['', '分表', '末端计量回路，参与结构与排名分析'],
        ['', '备用表', '不计入任何合计'],
        [],
        ['注意：总表与分表不可相加，否则重复计量；本系统只以总表口径作为总用量。'],
    ]
    return rows


def export_data():
    """原始数据全量导出"""
    ms = DB.meters()
    sheets = []
    rows = [['表计台账'], [],
            ['ID', '表编号', '表名称', '类型', '系统分类', '区域', '安装位置', '配电柜',
             '用途', '倍率', '计量层级', '主备', '折算系数', '水源', '启用']]
    for m in ms:
        lv = ('项目关口总表' if m.get('is_grid') else '一级总表' if m.get('is_total')
              else '二级总表' if m.get('is_sub_total') else '分表')
        rows.append([m['id'], m['code'], m['name'], '电' if m['kind'] == 'elec' else '水',
                     m['sys'], m['zone'], m['loc'], m['panel'], m['purpose'],
                     m['mult'], lv,
                     '备用' if m.get('use_status') == 'backup' else '主用',
                     m.get('factor') or 1,
                     '中水' if m.get('source') == 'reclaimed' else '自来水',
                     '是' if m['enabled'] else '否'])
    sheets.append(dict(title='表计台账', rows=rows,
                       widths=[6, 14, 30, 8, 12, 10, 20, 12, 30, 8, 13, 8, 9, 10, 8]))

    rd = DB.all_readings()
    mby = {m['id']: m for m in ms}
    rows = [['抄表记录'], [], ['表名称', '类型', '抄表日期', '表读数']]
    for (mid, dt), v in sorted(rd.items(), key=lambda x: (x[0][1], x[0][0])):
        m = mby.get(mid)
        if not m:
            continue
        rows.append([m['name'], '电' if m['kind'] == 'elec' else '水', dt, v])
    sheets.append(dict(title='抄表记录', rows=rows, widths=[30, 8, 14, 14]))
    return build_workbook(sheets)


def backup_json():
    ms = DB.meters()
    rd = DB.all_readings()
    d = dict(version=2, exported=DB.now(), meters=ms,
             readings=[dict(meter_id=k[0], rdate=k[1], val=v) for k, v in rd.items()],
             prices=DB.prices(),
             price_months={k: [dict(mk=x, unit_price=y) for x, y in DB.price_months(k).items()]
                           for k in ('elec', 'water')},
             groups=DB.groups(),
             dicts=[dict(kind=x['kind'], name=x['name'], sort=x['sort']) for x in _all_dicts()],
             meta={k: DB.meta_get(k) for k in ('company', 'project', 'copyright', 'title')})
    return json.dumps(d, ensure_ascii=False, indent=1).encode('utf-8')


def restore_json(content):
    d = json.loads(content)
    if 'meters' not in d:
        raise ValueError('缺少 meters 字段')
    c = DB.conn()
    c.execute('DELETE FROM reading')
    c.execute('DELETE FROM meter')
    c.commit()
    idmap = {}
    for m in d['meters']:
        new = DB.meter_save(dict(m, id=None))
        idmap[m['id']] = new
    for r in d.get('readings', []):
        m = idmap.get(r['meter_id'])
        if m:
            DB.reading_save(m, r['rdate'], r['val'])
    for k, v in (d.get('prices') or {}).items():
        DB.price_set(k, v)
    # 恢复分类字典与显示分组
    for x in d.get('dicts') or []:
        DB.dict_add(x.get('kind'), x.get('name'))
    if d.get('groups'):
        c2 = DB.conn()
        c2.execute('DELETE FROM grp')
        c2.commit(); c2.close()
        for g in d['groups']:
            DB.group_save(dict(g, id=None))
    # 恢复按月单价
    for kind in ('elec', 'water'):
        for pm in (d.get('price_months') or {}).get(kind, []) or []:
            DB.price_month_set(kind, pm.get('mk'), pm.get('unit_price'))
    for k, v in (d.get('meta') or {}).items():
        if k == 'copyright':
            continue          # 版权固定，不从备份覆盖
        DB.meta_set(k, v)
    return True
