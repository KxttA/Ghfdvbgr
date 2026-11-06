# -*- coding: utf-8 -*-
"""能耗分析引擎：同比、环比、结构、排名、异常检测、节能测算"""
import datetime
try:
    from server import db as DB
    from server.db import (agg_by, month_list, prev_month, same_month_last_year,
                           days_in_month, build_monthly, all_readings, meters, prices)
except ImportError:
    import db as DB
    from db import (agg_by, month_list, prev_month, same_month_last_year,
                    days_in_month, build_monthly, all_readings, meters, prices)


def _pct(cur, base):
    if base in (None, 0):
        return None
    return round((cur - base) / abs(base) * 100, 1)


def _f(v, n=1):
    if v is None:
        return '—'
    return ('%.' + str(n) + 'f') % v


def _skip(m, kind, include_reclaimed=False, zones=None):
    """统一筛选：水表默认剔除中水（单独统计）；zones 为区域白名单；备用表默认剔除"""
    if m['kind'] != kind:
        return True
    if kind == 'water' and m.get('source') == 'reclaimed' and not include_reclaimed:
        return True
    if m.get('use_status') == 'backup':
        return True
    if zones and (m.get('zone') or '') not in zones:
        return True
    return False


def _zone_out(m, zones):
    """区域筛选：zones 为空表示不筛选"""
    if not zones:
        return False
    return (m.get('zone') or '') not in zones


def _is_sum_meter(m):
    """是否总表类（一级总表或二级总表）——不进分表合计，避免重复计量"""
    return bool(m.get('is_total') or m.get('is_sub_total') or m.get('is_grid'))


class Analysis:
    def __init__(self):
        self.meters = meters()
        self.mby = {m['id']: m for m in self.meters}
        self.monthly = build_monthly(self.meters, all_readings(), prices())
        self._apply_prices()
        self.months = month_list(self.monthly)

    def _apply_prices(self):
        """把按月设置的单价套到各月金额上（电价每月可能不同）"""
        for kind in ('elec', 'water'):
            pm = DB.price_months(kind)
            if pm:
                base = prices().get(kind, 0)
                merged = {}
                for mid, v in self.monthly.items():
                    if v['meta']['kind'] != kind:
                        continue
                    for mk in v['months']:
                        merged.setdefault(mk, pm.get(mk, base))
                DB.apply_month_price(self.monthly, kind, merged)

    def price_of(self, kind, mk):
        """某月单价"""
        return DB.price_month(kind, mk)

    # ---------- 计量层级 ----------
    def grid_ids(self, kind):
        return [m['id'] for m in self.meters if m['kind'] == kind and m.get('is_grid')]

    def level1(self, kind, zones=None):
        """一级总表：is_total（低压配电柜进线柜 D101/D201、自来水总表）。
        国网关口表（is_grid）不算一级总表——它只用于核对，不参与总用量。"""
        l1 = [m['id'] for m in self.meters
              if m['kind'] == kind and m.get('is_total') and not _zone_out(m, zones)]
        if l1:
            return l1
        # 没设一级总表的项目，退回国网关口表兜底
        return [m['id'] for m in self.meters
                if m['kind'] == kind and m.get('is_grid') and not _zone_out(m, zones)]

    def level2(self, kind, zones=None):
        """二级总表（区域总表、BDEF区总水表等）"""
        return [m['id'] for m in self.meters
                if m['kind'] == kind and m.get('is_sub_total') and not m.get('is_grid')
                and m.get('use_status') != 'backup' and not _zone_out(m, zones)]

    def sub_ids(self, kind, zones=None, include_reclaimed=False):
        """分表（末端计量），排除总表/二级总表/国网/备用，水再排除中水"""
        out = []
        for m in self.meters:
            if _skip(m, kind, include_reclaimed, zones):
                continue
            if _is_sum_meter(m):
                continue
            out.append(m['id'])
        return out

    def _sum_ids(self, ids, months):
        res = {}
        for mid in ids:
            v = self.monthly.get(mid)
            if not v:
                continue
            f = float(v['meta'].get('factor') or 1)
            for mk, mv in v['months'].items():
                if mk in months:
                    r = res.setdefault(mk, dict(usage=0, cost=0))
                    r['usage'] += mv['usage'] * f
                    r['cost'] += mv['cost'] * f
        for k in res:
            res[k]['usage'] = round(res[k]['usage'], 1)
            res[k]['cost'] = round(res[k]['cost'], 1)
        return res

    # ---------- 通用 ----------
    def total(self, kind, months=None, zones=None, include_reclaimed=False, basis='meter'):
        """
        总用量月度合计。
        basis='meter'（默认）：按总表统计 —— 电=低压进线柜总表合计，水=自来水总表合计；
                              无总表数据的月份自动回落到分表合计，保证不漏数。
        basis='sub'：按分表合计（用于与总表对比）
        """
        months = months or self.months
        if basis == 'sub':
            return self._sum_ids(self.sub_ids(kind, zones, include_reclaimed), months)
        res = self._sum_ids(self.level1(kind, zones), months)
        # 总表无数据的月份 → 回落到 二级总表 → 分表合计
        sub = None
        lv2 = None
        for mk in months:
            if mk in res and res[mk]['usage']:
                continue
            if lv2 is None:
                lv2 = self._sum_ids(self.level2(kind, zones), months)
            if mk in lv2 and lv2[mk]['usage']:
                res[mk] = dict(usage=lv2[mk]['usage'], cost=lv2[mk]['cost'])
                continue
            if sub is None:
                sub = self._sum_ids(self.sub_ids(kind, zones, include_reclaimed), months)
            if mk in sub:
                res[mk] = dict(usage=sub[mk]['usage'], cost=sub[mk]['cost'])
        return res

    def total_month(self, kind, mk, use_total_meter=True):
        tiers = (['meter'], ['sub']) if use_total_meter else (['sub'],)
        for b in tiers:
            t = self.total(kind, [mk], basis=b).get(mk)
            if t and t['usage']:
                return t['usage']
        return 0

    # ---------- 本期 vs 上期（环比） ----------
    def compare(self, kind, mk, zones=None, include_reclaimed=False):
        """返回 {cur, prev, pmk, mom(pct), daily_cur, daily_prev, mom_daily(pct), yoy, ymk}"""
        tt = self.total(kind, zones=zones, include_reclaimed=include_reclaimed)
        cur = tt.get(mk, {}).get('usage', 0)
        pmk = prev_month(mk)
        prev = tt.get(pmk, {}).get('usage')
        ymk = same_month_last_year(mk)
        yo = tt.get(ymk, {}).get('usage')
        dc = round(cur / days_in_month(mk), 2) if cur else 0
        dp = round(prev / days_in_month(pmk), 2) if prev else None
        return dict(
            cur=round(cur, 1), cost=round(tt.get(mk, {}).get('cost', 0), 1),
            prev=round(prev, 1) if prev is not None else None, pmk=pmk,
            mom=_pct(cur, prev),
            daily_cur=dc, daily_prev=dp,
            mom_daily=(round((dc - dp) / dp * 100, 1) if dp else None),
            yoy=_pct(cur, yo), ymk=ymk, last_year=round(yo, 1) if yo is not None else None,
        )

    # ---------- 结构（按系统 / 区域） ----------
    def structure(self, kind, mk, key='sys', zones=None, include_reclaimed=False):
        """
        结构占比（系统分类 / 区域 / 表计分项）。
        口径：按**分表**统计（末端计量），合计可能大于总表（重复计量/未计量差），
        页面会同时展示总表值与该差额，两者口径不同不可直接相加。
        """
        prev_mk = prev_month(mk)
        yk = same_month_last_year(mk)
        st = {}
        for mid, v in self.monthly.items():
            m = v['meta']
            if _skip(m, kind, include_reclaimed, zones) or _is_sum_meter(m):
                continue
            lbl = m['name'] if key == 'cat' else (m.get(key) or '其他')
            d = st.setdefault(lbl, {'_id': mid} if key == 'cat' else {})
            f = float(m.get('factor') or 1)
            d[mk] = d.get(mk, 0) + v['months'].get(mk, {}).get('usage', 0) * f
            d[prev_mk] = d.get(prev_mk, 0) + v['months'].get(prev_mk, {}).get('usage', 0) * f
            d[yk] = d.get(yk, 0) + v['months'].get(yk, {}).get('usage', 0) * f

        items = []
        tot = sum(v.get(mk, 0) for v in st.values())
        for lbl, mm in st.items():
            u = mm.get(mk, 0)
            p = mm.get(prev_mk) or None
            y = mm.get(yk) or None
            it = dict(
                label=lbl, usage=round(u, 1), raw=round(u, 1),
                share=round(u / tot * 100, 1) if tot else 0,
                prev=round(p, 1) if p is not None else None,
                mom=_pct(u, p), yoy=_pct(u, y),
                cost_delta=0,
            )
            if key == 'cat':
                it['id'] = mm.get('_id')
            items.append(it)
        items.sort(key=lambda x: -x['usage'])
        unit = prices().get(kind, 0)
        for i in items:
            i['cost'] = round(i['usage'] * unit, 1)
            i['cost_delta'] = round((i['usage'] - (i['prev'] or i['usage'])) * unit, 1)
        return items

    def _meter_total(self, kind, mk, zones=None, include_reclaimed=False):
        """单月总表口径用量（电=一级总表合计，水=自来水总表）"""
        if not mk:
            return 0
        t = self.total(kind, [mk], zones=zones, include_reclaimed=include_reclaimed)
        return (t.get(mk) or {}).get('usage', 0) or 0

    def struct_trend(self, kind, months=None, zones=None, include_reclaimed=False):
        """分系统逐月用量（用于趋势对比图）：[{label, data:[...]}]"""
        months = months or self.months
        st = {}
        mset = set(months)
        for mid, v in self.monthly.items():
            m = v['meta']
            if _skip(m, kind, include_reclaimed, zones) or _is_sum_meter(m):
                continue
            lbl = m.get('sys') or '其他'
            mm = st.setdefault(lbl, {})
            f = float(m.get('factor') or 1)
            for mk in mset:
                mm[mk] = mm.get(mk, 0) + v['months'].get(mk, {}).get('usage', 0) * f
        out = []
        for lbl, mm in st.items():
            out.append(dict(label=lbl, data=[round(mm.get(m, 0), 1) for m in months],
                            total=round(sum(mm.values()), 1)))
        out.sort(key=lambda x: -x['total'])
        return out

    # ---------- 表计排名 ----------
    def top_meters(self, kind, mk, n=10, zones=None, include_reclaimed=False):
        rows = []
        for mid, v in self.monthly.items():
            m = v['meta']
            if _skip(m, kind, include_reclaimed, zones):
                continue
            if mk in v['months']:
                mv = v['months'][mk]
                pv = v['months'].get(prev_month(mk), {}).get('usage')
                rows.append(dict(id=mid, name=m['name'], sys=m['sys'], zone=m['zone'],
                                 usage=mv['usage'], cost=mv['cost'],
                                 prev=pv, mom=_pct(mv['usage'], pv),
                                 is_total=bool(m.get('is_total')),
                                 is_sub_total=bool(m.get('is_sub_total')),
                                 is_grid=bool(m.get('is_grid')),
                                 use_status=m.get('use_status') or 'main'))
        rows.sort(key=lambda x: -x['usage'])
        return rows[:n]

    # ---------- 异常检测 ----------
    def anomalies(self, kind, mk, zones=None, include_reclaimed=False):
        out = []
        unit = prices().get(kind, 0)
        for mid, v in self.monthly.items():
            m = v['meta']
            if _skip(m, kind, include_reclaimed, zones):
                continue
            name = '%s（%s）' % (m['name'], m['zone']) if m['zone'] else m['name']
            # 1) 读数倒走
            for iv in v['intervals']:
                if iv['backwards']:
                    out.append(dict(level='高', type='读数异常', meter=name, mid=mid,
                                    detail='%s → %s 读数由 %s 降至 %s，疑似换表或抄错'
                                           % (iv['start'], iv['end'], iv['prev_val'], iv['curr_val']),
                                    impact='本月用量可能失真'))
            # 2) 长期无变化（可能表坏）
            ivs = [i for i in v['intervals'] if not i['backwards']]
            if len(ivs) >= 3:
                last3 = ivs[-3:]
                if all(i['raw'] == 0 for i in last3) and not _is_sum_meter(m):
                    out.append(dict(level='中', type='疑似停走', meter=name, mid=mid,
                                    detail='连续 %d 次抄表读值完全不变（%s ~ %s）'
                                           % (len(last3), last3[0]['start'], last3[-1]['end']),
                                    impact='可能表计故障或该回路未投入使用'))
                elif all(i['raw'] > 0 and i['raw'] < 0.6 for i in last3) and not _is_sum_meter(m):
                    out.append(dict(level='低', type='用量极小', meter=name, mid=mid,
                                    detail='连续 %d 次用量不足 1 个单位' % len(last3),
                                    impact='可能空置或计量回路未负载'))
            # 3) 环比剧烈波动
            cur = v['months'].get(mk, {}).get('usage')
            pv = v['months'].get(prev_month(mk), {}).get('usage')
            if cur and pv and pv > 2:
                ch = (cur - pv) / pv * 100
                if ch > 80:
                    out.append(dict(level='高', type='用量剧增', meter=name, mid=mid,
                                    detail='%s 用量 %s，环比 +%.0f%%（上月 %s）'
                                           % (mk, _f(cur), ch, _f(pv)),
                                    impact='多耗约 %.0f 元，需核查是否新增设备或漏电漏水'
                                           % ((cur - pv) * unit)))
                elif ch < -60 and pv > 10:
                    out.append(dict(level='中', type='用量骤降', meter=name, mid=mid,
                                    detail='%s 用量 %s，环比 %.0f%%（上月 %s）'
                                           % (mk, _f(cur), ch, _f(pv)),
                                    impact='可能停用、检修或漏抄'))
            # 4) 超长间隔（漏抄）
            for iv in v['intervals']:
                if iv['days'] > 75 and iv['end'].startswith(mk[:7]):
                    out.append(dict(level='中', type='抄表间隔过长', meter=name, mid=mid,
                                    detail='%s ~ %s 间隔 %d 天，覆盖多个计费周期'
                                           % (iv['start'], iv['end'], iv['days']),
                                    impact='该段用量已按月均摊，但月度精度下降'))
            # 5) 换表清零（已自动按前一段日均估算，提示用户核对）
            for iv in v['intervals']:
                if iv.get('reset'):
                    out.append(dict(level='中', type='换表/表底清零', meter=name, mid=mid,
                                    detail='%s 读数已按新表底计；%s' % (iv['end'], iv.get('reset_note', '')),
                                    impact='该段用量按换表前日均估算，建议核对实际换表日期与表底'))
        order = {'高': 0, '中': 1, '低': 2}
        out.sort(key=lambda x: (order.get(x['level'], 3), x['type']))
        return out

    # ---------- 数据完整度 ----------
    def coverage(self, kind, mk, zones=None, include_reclaimed=False):
        tot = sum(1 for m in self.meters if m['kind'] == kind and m['enabled']
                  and not _is_sum_meter(m) and not _skip(m, kind, include_reclaimed, zones))
        has = sum(1 for mid, v in self.monthly.items()
                  if v['meta']['kind'] == kind and mk in v['months']
                  and not _is_sum_meter(v['meta'])
                  and not _skip(v['meta'], kind, include_reclaimed, zones))
        return dict(total=tot, covered=has,
                    pct=round(has / tot * 100, 1) if tot else 0)

    # ---------- 节能测算 ----------
    def savings(self, kind, mk):
        """
        基于结构占比与环比变化，给出可执行的节能建议及测算
        """
        items = self.structure(kind, mk)
        unit = prices().get(kind, 0)
        unitword = '度' if kind == 'elec' else '吨'
        res = []
        tot = sum(i['usage'] for i in items) or 1

        # 规则库：(匹配系统名, 标题, 判定条件, 建议, 节降比例)
        if kind == 'elec':
            rules = [
                ('照明', '公区/商铺照明分时管控',
                 lambda i: i['share'] >= 15,
                 '照明占总用电 {share}%。建议：①地下车库、通道照明按 50% 隔灯开启并加装人体感应；'
                 '②营业结束后 23:00—07:00 仅保留应急照明；③逐步替换为 LED 并加装智能时控。',
                 0.18),
                ('空调用电', '空调温度与运行时段优化',
                 lambda i: i['share'] >= 10,
                 '空调占总用电 {share}%，且受季节影响最大。建议：夏季设定不低于 26℃、冬季不高于 20℃；'
                 '过渡季（4—5月、10—11月）关闭主机改全新风；加装变频与分区控制。',
                 0.12),
                ('景观照明', '景观照明定时压缩',
                 lambda i: i['share'] >= 5,
                 '景观照明占 {share}%。建议：与市政亮化时间同步，非节假日 22:00 后关闭 50%；'
                 '更换高光效灯具，实测照度达标即可。',
                 0.25),
                ('电梯扶梯', '扶梯变频与空载停驶',
                 lambda i: i['share'] >= 8,
                 '电梯扶梯占 {share}%。建议：扶梯加装变频器，无人时低速运行；夜间保留 1 台直梯。',
                 0.15),
                ('排水泵', '排水泵运行策略核查',
                 lambda i: i['share'] >= 3 and i.get('mom') and i['mom'] > 30,
                 '排水泵用电环比上升 {mom}%。建议：核查液位控制器设定与泵体是否有空转、'
                 '集水坑是否有渗漏持续补水。',
                 0.10),
                ('临电转供', '临电转供计量与结算',
                 lambda i: i['share'] >= 10,
                 '临电转供占 {share}%，属代收代付性质。建议：单独装表计量并与施工单位按月结算，'
                 '避免公区电费被摊薄。',
                 0.0),
            ]
        else:
            rules = [
                ('卫生间', '卫生间用水器具改造',
                 lambda i: i['share'] >= 25,
                 '卫生间用水占 {share}%。建议：更换感应式水嘴与节水型蹲便器（3/6L 双档）；'
                 '加装减压阀控制水压；夜间巡检跑冒滴漏。',
                 0.15),
                ('绿化灌溉', '绿化灌溉方式优化',
                 lambda i: i['share'] >= 20,
                 '绿化灌溉占 {share}%。建议：改滴灌/微喷替代漫灌；按土壤湿度或降雨联动控制；'
                 '避开正午高温时段，改早晚浇水；优先使用中水。',
                 0.20),
                ('绿化雾森', '雾森运行时段压缩',
                 lambda i: i['share'] >= 10,
                 '雾森占 {share}%。建议：仅在客流高峰（11:00—13:00、17:00—19:00）开启；'
                 '加装湿度联锁，雨天自动停机。',
                 0.30),
                ('水景', '水景循环与补水控制',
                 lambda i: i['share'] >= 5,
                 '水景补水占 {share}%。建议：检查水池渗漏；循环水加装过滤减少换水频次；'
                 '非开放时段停泵。',
                 0.20),
                ('消防补水', '消防管网保压查漏',
                 lambda i: i['share'] >= 10,
                 '消防补水占 {share}%，正常应为偶发补水。若持续偏高，说明管网或水池存在渗漏，'
                 '建议做保压试验定位漏点。',
                 0.30),
            ]

        for sysname, title, cond, tpl, ratio in rules:
            for it in items:
                if sysname in it['label'] and cond(it):
                    save_u = it['usage'] * ratio
                    res.append(dict(
                        title=title,
                        sys=it['label'],
                        basis='本月 %s 占总用量 %.1f%%，用电/水量 %s %s'
                              % (it['label'], it['share'], _f(it['usage']), unitword),
                        advice=tpl.format(share='%.1f%%' % it['share'],
                                          mom=('%+.0f%%' % it['mom']) if it.get('mom') is not None else '—'),
                        save_usage=round(save_u, 1),
                        save_cost=round(save_u * unit, 0),
                        ratio=int(ratio * 100),
                    ))
                    break

        # 全局建议
        tot_cost = sum(i['usage'] for i in items) * unit
        if res:
            su = sum(r['save_usage'] for r in res)
            res.append(dict(title='合计测算', sys='—',
                            basis='以上措施全部落实到位的理论节降额度',
                            advice='按各项节降比例保守估算，实际效果需结合实际工况验证，建议先试点 1—3 个月',
                            save_usage=round(su, 1),
                            save_cost=round(sum(r['save_cost'] for r in res), 0),
                            ratio=round(su / tot * 100, 1) if tot else 0,
                            is_sum=True))
        return res

    # ---------- 分组趋势（A+D+北通道 / E+F 等，自由设置） ----------
    def group_series(self, kind, months=None, groups=None):
        """每个分组在各月的用量/费用，返回 [{id,name,color,months:[{month,usage,cost}],total}]"""
        months = months or self.months
        grps = groups if groups is not None else DB.groups()
        out = []
        for g in grps:
            if g.get('kind') not in ('both', kind):
                continue
            zl = set(g.get('zones') or [])
            sl = set(g.get('sys') or [])
            data = {mk: dict(usage=0.0, cost=0.0) for mk in months}
            for mid, v in self.monthly.items():
                m = v['meta']
                if m['kind'] != kind or m.get('is_total'):
                    continue
                if zl and (m.get('zone') or '') not in zl:
                    continue
                if sl and (m.get('sys') or '') not in sl:
                    continue
                for mk, mv in v['months'].items():
                    if mk not in data:
                        continue
                    data[mk]['usage'] += mv['usage']
                    data[mk]['cost'] += mv['cost']
            out.append(dict(id=g.get('id'), name=g['name'], color=g.get('color') or '',
                            zones=g.get('zones') or [], sys=g.get('sys') or [],
                            months=[dict(month=mk, usage=round(data[mk]['usage'], 1),
                                         cost=round(data[mk]['cost'], 1))
                                    for mk in months],
                            total=round(sum(data[mk]['usage'] for mk in months), 1)))
        return out

    # ---------- 抄表质量（可按电/水与年月筛选） ----------
    def quality(self, kind=None, mk=None, zone=None):
        """所有表的抄表质量画像；kind 过滤电/水；mk（YYYY-MM 或 YYYY）只统计该期；zone 过滤区域/系统"""
        out = []
        for mid, v in self.monthly.items():
            m = v['meta']
            if kind and m['kind'] != kind:
                continue
            if zone and zone not in (m.get('zone') or '', m.get('sys') or '', m.get('loc') or ''):
                continue
            ivs = v['intervals']
            if mk:
                span = mk[:7] if len(mk) > 4 else mk
                ivs = [i for i in ivs if i['end'].startswith(span) or i['start'].startswith(span)]
            if not ivs:
                if not mk:
                    out.append(dict(id=mid, name=m['name'], kind=m['kind'], sys=m['sys'],
                                    zone=m['zone'], cnt=0, avg_gap=0, max_gap=0, status='无数据'))
                continue
            gaps = [i['days'] for i in ivs]
            avg = sum(gaps) / len(gaps)
            mx = max(gaps)
            status = '正常'
            if mx > 45:
                status = '间隔过长'
            if len([i for i in ivs if i['backwards']]):
                status = '读数倒走'
            out.append(dict(id=mid, name=m['name'], kind=m['kind'], sys=m['sys'], zone=m['zone'],
                            cnt=len(ivs), avg_gap=round(avg, 0), max_gap=mx, status=status))
        out.sort(key=lambda x: (x['kind'], -x['max_gap']))
        return out

    # ---------- 年汇总（年度每月总表） ----------
    def year_summary(self, kind, year, zones=None, include_reclaimed=False):
        tt = self.total(kind, zones=zones, include_reclaimed=include_reclaimed)
        rows = []
        for mo in range(1, 13):
            mk = '%04d-%02d' % (int(year), mo)
            v = tt.get(mk)
            days = days_in_month(mk)
            rows.append(dict(month=mk,
                             usage=round(v['usage'], 1) if v else None,
                             cost=round(v['cost'], 1) if v else None,
                             price=self.price_of(kind, mk),
                             days=days,
                             daily=round(v['usage'] / days, 2) if v else None))
        have = [r for r in rows if r['usage'] is not None]
        tot_u = round(sum(r['usage'] for r in have), 1)
        tot_c = round(sum(r['cost'] for r in have), 1)
        return dict(year=int(year), rows=rows, total_usage=tot_u, total_cost=tot_c,
                    months_count=len(have),
                    avg_month=round(tot_u / len(have), 1) if have else 0,
                    max_month=max(have, key=lambda r: r['usage'])['month'] if have else '',
                    min_month=min(have, key=lambda r: r['usage'])['month'] if have else '')

    # ---------- 中水（绿化灌溉等）单独核算与分摊 ----------
    def reclaimed(self, months=None, zones=None):
        """
        中水核算：
        - 补水 = sys=='中水补水' 的表（接自来水）的用量
        - 出水 = source=='reclaimed' 的中水用水表用量（按区域）
        - 分摊：某区中水用量 ÷ 中水表总量 × 中水补水量（该区实际消耗的自来水）
        - 损耗 = 补水 - 出水合计
        返回 {months, supply, reclaimed, loss, loss_pct, zones:[{zone,usage,share,alloc}], by_month}
        """
        months = months or self.months
        supply_ids = [m['id'] for m in self.meters
                      if m['kind'] == 'water' and m.get('sys') == '中水补水' and not m.get('is_total')]
        rec = {}       # mid -> meta
        for m in self.meters:
            if m['kind'] == 'water' and m.get('source') == 'reclaimed':
                if zones and (m.get('zone') or '') not in zones:
                    continue
                rec[m['id']] = m
        sup_by_month, rec_by_month = {}, {}
        zone_tot = {}
        for mid, v in self.monthly.items():
            m = v['meta']
            for mk, mv in v['months'].items():
                if mk not in months:
                    continue
                if mid in supply_ids or (m.get('sys') == '中水补水' and not m.get('is_total')):
                    sup_by_month[mk] = sup_by_month.get(mk, 0) + mv['usage']
                elif mid in rec:
                    rec_by_month[mk] = rec_by_month.get(mk, 0) + mv['usage']
                    z = m.get('zone') or '未分区'
                    zone_tot[z] = zone_tot.get(z, 0) + mv['usage']
        tot_rec = sum(zone_tot.values())
        tot_sup = sum(sup_by_month.values())
        zl = []
        for z, u in sorted(zone_tot.items(), key=lambda x: -x[1]):
            zl.append(dict(zone=z, usage=round(u, 1),
                           share=round(u / tot_rec * 100, 1) if tot_rec else 0,
                           alloc=round(u / tot_rec * tot_sup, 1) if tot_rec else 0))
        loss = round(tot_sup - tot_rec, 1)
        return dict(months=months,
                    supply=round(tot_sup, 1), reclaimed=round(tot_rec, 1), loss=loss,
                    loss_pct=round(loss / tot_sup * 100, 1) if tot_sup else 0,
                    zones=zl,
                    by_month=[dict(month=mk,
                                   supply=round(sup_by_month.get(mk, 0), 1),
                                   reclaimed=round(rec_by_month.get(mk, 0), 1),
                                   loss=round(sup_by_month.get(mk, 0) - rec_by_month.get(mk, 0), 1))
                              for mk in months])

    # ---------- 损耗分析（三级总表比对 + 国网抄表核对） ----------
    def loss(self, kind='elec', months=None, zones=None):
        """
        三类总表横向比对：
          G  国网关口表（is_grid，高压结算表）
          L1 一级总表（is_total，低压配电柜进线柜 D101/D201）
          SUB 分表合计（末端回路，不含总表/备用/中水）
        电：
          变损线损（一次）= G − L1
          未计量（二次）  = L1 − SUB
          总损耗          = G − SUB ；总损耗率 = 总损耗 ÷ G
          国网误差率      = (G − L1) ÷ G   即自抄一级表与国网的偏差
        水：
          损耗 = 自来水总表 − 自来水分表合计（中水排除）
        每行均带环比/同比（用量、差值、损耗率三条线都有）
        """
        months = months or self.months
        gids = list(self.grid_ids(kind))
        # 一级总表 = is_total（低压进线柜），与国网关口表严格分开；
        # 若项目没设关口表，则 level1 会回落到 is_total，这里也兼容
        l1ids = [m['id'] for m in self.meters
                 if m['kind'] == kind and m.get('is_total')
                 and m.get('use_status') != 'backup' and not _zone_out(m, zones)]
        if not l1ids:
            l1ids = list(self.level1(kind, zones))
        lv2 = self.level2(kind, zones)
        subs = self.sub_ids(kind, zones)
        has_grid = bool(gids)

        # 国网结算量：优先取「国网抄表」录入的结算量（grid_reading），其次按关口表抄表读数
        gr = {x['mk']: x for x in DB.grid_readings(kind)}
        L1 = self._sum_ids(l1ids, months)
        Gm = self._sum_ids(gids, months)

        def _grid_usage(mk):
            gv = gr.get(mk) or {}
            if gv.get('usage') is not None:
                return round(gv['usage'], 1)
            return round((Gm.get(mk, {}) or {}).get('usage', 0) or 0, 1)
        L2 = self._sum_ids(lv2, months)
        SUB = self._sum_ids(subs, months)

        def _u(d, mk):
            return round((d.get(mk, {}) or {}).get('usage', 0) or 0, 1)

        rows = []
        for mk in months:
            g = _grid_usage(mk)
            l1 = _u(L1, mk)
            l2 = _u(L2, mk)
            sb = _u(SUB, mk)
            rows.append(dict(
                month=mk,
                g=g, l1=l1, l2=l2, sub=sb,
                # 三类总表并列（报表/页面直接用这三列比）
                d_g_l1=round(g - l1, 1) if g else None,          # 国网 − 一级表
                d_g_l1_pct=round((g - l1) / g * 100, 2) if g else None,
                d_l1_sub=round(l1 - sb, 1) if l1 else None,      # 一级表 − 分表
                d_l1_sub_pct=round((l1 - sb) / l1 * 100, 2) if l1 else None,
                d_g_sub=round(g - sb, 1) if g else None,         # 国网 − 分表（总损耗）
                d_g_sub_pct=round((g - sb) / g * 100, 2) if g else None,
                # 兼容旧字段
                grid=g, lv=l1, sub2=l2, total=l1 or g, sub_total=l2 or sb,
                loss1=round(g - l1, 1) if (g and l1) else None,
                loss2=round(l1 - (l2 or sb), 1) if l1 else None,
                loss_all=round((g or l1) - (l2 or sb), 1) if (g or l1) else None,
                rate=round(((g or l1) - (l2 or sb)) / (g or l1) * 100, 1) if (g or l1) else None,
                lv_rate=round((l1 - (l2 or sb)) / l1 * 100, 1) if l1 else None))

        def _pair(key, rows, i):
            """取上期/去年同期同字段，返回 (mom, yoy)"""
            cur = rows[i].get(key)
            pv = rows[i - 1].get(key) if i else None
            yv = next((x.get(key) for x in rows
                       if x['month'] == same_month_last_year(rows[i]['month'])), None)
            return (_pct(cur, pv) if (pv not in (None, 0) and cur is not None) else None,
                    _pct(cur, yv) if (yv not in (None, 0) and cur is not None) else None)

        for i, r in enumerate(rows):
            for k in ('g', 'l1', 'sub', 'd_g_l1', 'd_l1_sub', 'd_g_sub', 'rate', 'loss_all'):
                m, y = _pair(k, rows, i)
                r[k + '_mom'] = m
                r[k + '_yoy'] = y
            # 兼容旧字段名
            r['mom'] = r['loss_all_mom']
            r['yoy'] = r['loss_all_yoy']

        # 国网抄表核对（国网系统给的结算量 vs 自抄一级表）
        chk = []
        for r in rows:
            gv = gr.get(r['month']) or {}
            gu = gv.get('usage')
            mine = r['l1'] or r['g']
            diff = round(gu - mine, 1) if gu is not None else None
            chk.append(dict(month=r['month'], grid_usage=gu, mine=mine, diff=diff,
                            gap=round(diff / gu * 100, 2) if (gu and diff is not None) else None,
                            rdate=gv.get('rdate') or '', fee=gv.get('fee'),
                            note=gv.get('note') or ''))
        for i, c in enumerate(chk):
            pv = chk[i - 1] if i else None
            yv = next((x for x in chk if x['month'] == same_month_last_year(c['month'])), None)
            c['mom'] = _pct(c['diff'], pv['diff']) if (pv and c['diff'] is not None and pv['diff'] is not None) else None
            c['yoy'] = _pct(c['diff'], yv['diff']) if (yv and c['diff'] is not None and yv['diff'] is not None) else None

        # 各级明细（每块表各月用量 + 环比 + 同比）
        detail = []
        for mid in (gids + l1ids):
            v = self.monthly.get(mid) or {}
            ms_ = v.get('months', {})
            series = [round((ms_.get(mk, {}) or {}).get('usage', 0) or 0, 1) for mk in months]
            mom, yoy = [], []
            for j, mk in enumerate(months):
                pv = (ms_.get(months[j - 1], {}) or {}).get('usage') if j else None
                yvv = (ms_.get(same_month_last_year(mk), {}) or {}).get('usage')
                mom.append(_pct(series[j], pv) if pv else None)
                yoy.append(_pct(series[j], yvv) if yvv else None)
            detail.append(dict(id=mid, name=self.mby[mid]['name'],
                               grid=bool(self.mby[mid].get('is_grid')),
                               data=series, mom=mom, yoy=yoy,
                               total=round(sum(series), 1)))

        return dict(kind=kind, rows=rows, months=months, check=chk, detail=detail,
                    has_grid=has_grid,
                    grid=[self.mby[i]['name'] for i in gids],
                    level1=[self.mby[i]['name'] for i in l1ids],
                    level2=[self.mby[i]['name'] for i in lv2],
                    sub_count=len(subs))

    def years(self):
        return sorted({m[:4] for m in self.months})
