# -*- coding: utf-8 -*-
"""能耗分析系统 - 数据层：建库、表计台账、抄表记录、导入、能耗计算"""
import os, sys, sqlite3, json, datetime, calendar, hashlib, bisect

APP_DIR = os.path.dirname(os.path.abspath(__file__))


def data_dir():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(APP_DIR)
    d = os.path.join(base, 'data')
    os.makedirs(d, exist_ok=True)
    return d


DB_PATH = os.environ.get('ENERGY_DB') or os.path.join(data_dir(), 'energy.db')

SCHEMA = '''
CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT);

CREATE TABLE IF NOT EXISTS meter(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  code TEXT,                       -- 表编号（同一天可能有多个，允许重复）
  name TEXT NOT NULL,              -- 位置/表名
  kind TEXT NOT NULL,              -- elec / water
  sys TEXT NOT NULL,               -- 系统分类（照明、消防、电梯…）
  zone TEXT DEFAULT '',            -- 区域（A区、B区、D区、下沉广场…）
  loc TEXT DEFAULT '',             -- 详细位置/配电室/水表井
  panel TEXT DEFAULT '',           -- 所属配电柜
  purpose TEXT DEFAULT '',         -- 用途描述（原表"控制区域"）
  mult REAL DEFAULT 1,             -- 互感器倍率
  unit TEXT DEFAULT '',            -- 计量单位
  is_total INTEGER DEFAULT 0,      -- 总表（项目/区域一级计量，如 D101/D201 低压进线柜）
  is_sub_total INTEGER DEFAULT 0,  -- 二级总表（区域总表、BDEF区总水表等，隶属某个总表）
  parent_id INTEGER DEFAULT 0,     -- 所属上级总表 id（0=直接挂项目总表）
  use_status TEXT DEFAULT 'main',  -- 主备：main=主用 / backup=备用（备用表不计入合计）
  factor REAL DEFAULT 1,           -- 损耗/误差折算系数（实测校验用，通常 1）
  source TEXT DEFAULT 'tap',       -- 水源：tap=自来水 / reclaimed=中水（仅水表）
  is_grid INTEGER DEFAULT 0,       -- 项目关口总表（国网/高压，用于总用电量与损耗）
  enabled INTEGER DEFAULT 1,
  sort INTEGER DEFAULT 0,
  created TEXT
);

CREATE TABLE IF NOT EXISTS reading(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  meter_id INTEGER NOT NULL,
  rdate TEXT NOT NULL,             -- YYYY-MM-DD 抄表日期
  val REAL,                        -- 表读值（原始读数），NULL=漏抄
  note TEXT DEFAULT '',
  created TEXT,
  UNIQUE(meter_id, rdate)
);
CREATE INDEX IF NOT EXISTS ix_reading ON reading(meter_id, rdate);

CREATE TABLE IF NOT EXISTS price(
  kind TEXT PRIMARY KEY,           -- elec / water（默认单价）
  unit_price REAL DEFAULT 0,       -- 元/度 或 元/吨
  updated TEXT
);

-- 按月单价（电价每月不同，按账期月份生效）
CREATE TABLE IF NOT EXISTS price_m(
  kind TEXT NOT NULL,
  mk TEXT NOT NULL,                -- YYYY-MM
  unit_price REAL NOT NULL,
  updated TEXT,
  PRIMARY KEY(kind, mk)
);

-- 分类字典（系统分类 / 区域，用户可自行增加）
CREATE TABLE IF NOT EXISTS dict(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kind TEXT NOT NULL,              -- sys_elec / sys_water / zone
  name TEXT NOT NULL,
  sort INTEGER DEFAULT 0,
  created TEXT,
  UNIQUE(kind, name)
);

-- 显示分组（如 A+D+北通道 一组、E+F 一组，可自由设置），用于趋势对比
CREATE TABLE IF NOT EXISTS grp(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  zone_csv TEXT DEFAULT '',        -- 区域（逗号分隔）
  sys_csv TEXT DEFAULT '',         -- 系统（逗号分隔，空=全部）
  kind TEXT DEFAULT 'both',        -- elec / water / both
  color TEXT DEFAULT '',
  sort INTEGER DEFAULT 0,
  created TEXT
);

-- 国网抄表（供电公司每月出具的表底数与电量，用于和自抄比对）
CREATE TABLE IF NOT EXISTS grid_reading(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kind TEXT NOT NULL DEFAULT 'elec',   -- elec / water
  meter_no TEXT DEFAULT '',            -- 国网户号/表号
  mk TEXT NOT NULL,                    -- 账期 YYYY-MM
  rdate TEXT DEFAULT '',               -- 国网抄表日
  val REAL,                            -- 国网表底数
  usage REAL,                          -- 国网结算电量/水量
  fee REAL,                            -- 国网电费/水费（可选）
  note TEXT DEFAULT '',
  updated TEXT,
  UNIQUE(kind, mk)
);

CREATE TABLE IF NOT EXISTS log(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT, user TEXT, act TEXT, detail TEXT, ip TEXT
);
'''

# ---------- 系统分类（电力） ----------
ELEC_SYS = ['照明', '商铺照明', '景观照明', '消防', '电梯扶梯', '空调用电',
            '排水泵', '动力排风', '安防', '临电转供', '其他']
# ---------- 系统分类（水） ----------
WATER_SYS = ['卫生间', '绿化灌溉', '绿化雾森', '水景', '消防补水', '中水补水',
             '设备间', '其他']
# ---------- 区域 ----------
ZONES = ['A区', 'B区', 'C区', 'D区', 'E区', 'F区', '北通道', '下沉广场', '地上广场', '其他']

DEFAULT_PRICE = {'elec': 0.85, 'water': 4.5}


def conn():
    c = sqlite3.connect(DB_PATH, timeout=15)
    c.row_factory = sqlite3.Row
    c.execute('PRAGMA journal_mode=WAL')
    return c


def now():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def today():
    return datetime.date.today().isoformat()


def init():
    c = conn()
    c.executescript(SCHEMA)
    # 旧库字段迁移
    for col, ddl in (('source', "ALTER TABLE meter ADD COLUMN source TEXT DEFAULT 'tap'"),
                     ('is_grid', 'ALTER TABLE meter ADD COLUMN is_grid INTEGER DEFAULT 0'),
                     ('is_sub_total', 'ALTER TABLE meter ADD COLUMN is_sub_total INTEGER DEFAULT 0'),
                     ('parent_id', 'ALTER TABLE meter ADD COLUMN parent_id INTEGER DEFAULT 0'),
                     ('use_status', "ALTER TABLE meter ADD COLUMN use_status TEXT DEFAULT 'main'"),
                     ('factor', 'ALTER TABLE meter ADD COLUMN factor REAL DEFAULT 1')):
        try:
            c.execute(ddl)
        except Exception:
            pass
    # 历史数据自动标记：名称含"中水"→中水
    c.execute("UPDATE meter SET source='reclaimed' WHERE kind='water' AND (name LIKE '%中水%' AND name NOT LIKE '%补水%')")
    # 名称含"备用"的历史表 → 主备=备用（不计入合计，避免主备重复计量）
    c.execute("UPDATE meter SET use_status='backup' WHERE name LIKE '%备用%' OR purpose LIKE '%备用%'")
    c.execute("UPDATE meter SET factor=1 WHERE factor IS NULL OR factor=0")
    c.commit()
    for k, v in DEFAULT_PRICE.items():
        c.execute('INSERT OR IGNORE INTO price(kind,unit_price,updated) VALUES(?,?,?)',
                  (k, v, now()))
    # 分类字典初始化（首次）
    for dk, names in (('sys_elec', ELEC_SYS), ('sys_water', WATER_SYS), ('zone', ZONES)):
        n = c.execute('SELECT COUNT(*) n FROM dict WHERE kind=?', (dk,)).fetchone()['n']
        if not n:
            for i, nm in enumerate(names):
                c.execute('INSERT OR IGNORE INTO dict(kind,name,sort,created) VALUES(?,?,?,?)',
                          (dk, nm, i, now()))
    # 默认显示分组（可在系统设置里改）
    n = c.execute('SELECT COUNT(*) n FROM grp').fetchone()['n']
    if not n:
        for i, (nm, zc, co) in enumerate([
                ('A区·D区·北通道', 'A区,D区,北通道', '#2f6fd0'),
                ('E区·F区', 'E区,F区', '#0ea5a4')]):
            c.execute('INSERT INTO grp(name,zone_csv,sys_csv,kind,color,sort,created) '
                      'VALUES(?,?,?,?,?,?,?)', (nm, zc, '', 'both', co, i, now()))
    c.commit()
    c.close()


def dict_get(kind):
    """按类型取字典项名称列表（kind: sys_elec / sys_water / zone）"""
    c = conn()
    r = [x['name'] for x in c.execute(
        'SELECT name FROM dict WHERE kind=? ORDER BY sort,id', (kind,))]
    c.close()
    if r:
        return r
    return {'sys_elec': ELEC_SYS, 'sys_water': WATER_SYS, 'zone': ZONES}.get(kind, [])


def dict_add(kind, name):
    name = (name or '').strip()
    if not name:
        return False
    c = conn()
    c.execute('INSERT OR IGNORE INTO dict(kind,name,sort,created) VALUES(?,?,?,?)',
              (kind, name, 999, now()))
    c.commit(); c.close()
    return True


def dict_del(did):
    """删除字典项；返回 (ok, msg)"""
    c = conn()
    r = c.execute('SELECT * FROM dict WHERE id=?', (did,)).fetchone()
    if not r:
        c.close()
        return False, '字典项不存在'
    kind, name = r['kind'], r['name']
    # 有表计在用则禁止删除，避免分析口径丢失
    if kind == 'zone':
        used = c.execute('SELECT COUNT(*) n FROM meter WHERE zone=?', (name,)).fetchone()['n']
    else:
        k2 = 'elec' if kind == 'sys_elec' else 'water'
        used = c.execute('SELECT COUNT(*) n FROM meter WHERE sys=? AND kind=?',
                         (name, k2)).fetchone()['n']
    if used:
        c.close()
        return False, '还有 %d 块表在使用「%s」，请先改掉这些表的分类' % (used, name)
    c.execute('DELETE FROM dict WHERE id=?', (did,))
    c.commit(); c.close()
    return True, ''


def dict_rename(did, name):
    name = (name or '').strip()
    if not name:
        return False, '名称不能为空'
    c = conn()
    r = c.execute('SELECT * FROM dict WHERE id=?', (did,)).fetchone()
    if not r:
        c.close()
        return False, '字典项不存在'
    old, kind = r['name'], r['kind']
    c.execute('UPDATE dict SET name=? WHERE id=?', (name, did))
    # 同步把表计上的旧分类名改成新名
    if kind == 'zone':
        c.execute('UPDATE meter SET zone=? WHERE zone=?', (name, old))
    else:
        k2 = 'elec' if kind == 'sys_elec' else 'water'
        c.execute('UPDATE meter SET sys=? WHERE sys=? AND kind=?', (name, old, k2))
    c.commit(); c.close()
    return True, ''


# ================= 显示分组 =================
def groups():
    c = conn()
    r = [dict(x) for x in c.execute('SELECT * FROM grp ORDER BY sort,id')]
    c.close()
    for x in r:
        x['zones'] = [z for z in (x['zone_csv'] or '').split(',') if z.strip()]
        x['sys'] = [s for s in (x['sys_csv'] or '').split(',') if s.strip()]
    return r


def group_save(d):
    c = conn()
    zones = d.get('zones') or []
    sysl = d.get('sys') or []
    if isinstance(zones, str):
        zones = [z for z in zones.split(',') if z.strip()]
    if isinstance(sysl, str):
        sysl = [s for s in sysl.split(',') if s.strip()]
    vals = [d.get('name') or '未命名分组', ','.join(zones), ','.join(sysl),
            d.get('kind') or 'both', d.get('color') or '', int(d.get('sort') or 0)]
    if d.get('id'):
        c.execute('UPDATE grp SET name=?,zone_csv=?,sys_csv=?,kind=?,color=?,sort=? WHERE id=?',
                  vals + [d['id']])
        gid = d['id']
    else:
        c.execute('INSERT INTO grp(name,zone_csv,sys_csv,kind,color,sort,created) '
                  'VALUES(?,?,?,?,?,?,?)', vals + [now()])
        gid = c.execute('SELECT last_insert_rowid() id').fetchone()[0]
    c.commit(); c.close()
    return gid


def group_delete(gid):
    c = conn()
    c.execute('DELETE FROM grp WHERE id=?', (gid,))
    c.commit(); c.close()


def meta_get(k, default=''):
    c = conn()
    r = c.execute('SELECT v FROM meta WHERE k=?', (k,)).fetchone()
    c.close()
    return r['v'] if r else default


def meta_set(k, v):
    c = conn()
    c.execute('INSERT INTO meta(k,v) VALUES(?,?) ON CONFLICT(k) DO UPDATE SET v=excluded.v', (k, str(v)))
    c.commit(); c.close()


# ================= 密码 =================
def _h(pwd, salt):
    return hashlib.sha256((pwd + salt).encode('utf-8')).hexdigest()


def has_admin():
    return bool(meta_get('admin_pwd'))


def verify_pwd(pwd, h=None, salt=None):
    h = h or meta_get('admin_pwd')
    salt = salt or meta_get('admin_salt')
    if not h:
        return True
    return _h(pwd, salt) == h


def set_pwd(pwd):
    salt = os.urandom(8).hex()
    meta_set('admin_salt', salt)
    meta_set('admin_pwd', _h(pwd, salt))


# ================= 表计台账 =================
def meters(kind=None):
    c = conn()
    q = 'SELECT * FROM meter'
    args = []
    if kind:
        q += ' WHERE kind=?'
        args.append(kind)
    q += ' ORDER BY kind, sys, zone, sort, id'
    r = [dict(x) for x in c.execute(q, args)]
    c.close()
    return r


def meter_save(d):
    c = conn()
    f = ['code', 'name', 'kind', 'sys', 'zone', 'loc', 'panel', 'purpose',
         'mult', 'unit', 'is_total', 'is_sub_total', 'parent_id', 'use_status',
         'factor', 'source', 'is_grid', 'enabled', 'sort']
    vals = [d.get(k) for k in f]
    if not vals[2] and d.get('id'):
        r = c.execute('SELECT kind FROM meter WHERE id=?', (d['id'],)).fetchone()
        vals[2] = r['kind'] if r else None
    if not vals[2]:
        vals[2] = 'elec'
    kind = vals[2]
    vals[8] = float(d.get('mult') or 1)                                  # mult
    vals[10] = 1 if d.get('is_total') in (1, True, '1') else 0           # is_total
    vals[11] = 1 if d.get('is_sub_total') in (1, True, '1') else 0       # is_sub_total
    try:
        vals[12] = int(d.get('parent_id') or 0)                          # parent_id
    except (TypeError, ValueError):
        vals[12] = 0
    vals[13] = 'backup' if d.get('use_status') == 'backup' else 'main'   # use_status
    try:
        vals[14] = float(d.get('factor') or 1)                           # factor
    except (TypeError, ValueError):
        vals[14] = 1
    src = d.get('source') or 'tap'
    if kind == 'water' and d.get('sys') == '中水补水':
        src = 'tap'
    vals[15] = 'reclaimed' if (kind == 'water' and src == 'reclaimed') else 'tap'
    vals[16] = 1 if d.get('is_grid') in (1, True, '1') else 0            # is_grid
    vals[17] = 1 if d.get('enabled', 1) in (1, True, '1', 'true') else 0
    vals[18] = int(d.get('sort') or 0)
    # 总表/二级总表互斥：国网关口表不算总表
    if vals[16]:
        vals[10] = 0; vals[11] = 0
    if vals[10]:
        vals[11] = 0
    if d.get('id'):
        sets = ','.join('%s=?' % k for k in f)
        c.execute('UPDATE meter SET %s WHERE id=?' % sets, vals + [d['id']])
        mid = d['id']
    else:
        c.execute('INSERT INTO meter(%s,created) VALUES(%s,?)' % (','.join(f), ','.join('?' * len(f))),
                  vals + [now()])
        mid = c.execute('SELECT last_insert_rowid() id').fetchone()[0]
    c.commit(); c.close()
    return mid


def grid_readings(kind=None):
    """国网抄表记录 {mk: {...}}"""
    c = conn()
    q = 'SELECT * FROM grid_reading'
    a = []
    if kind:
        q += ' WHERE kind=?'; a.append(kind)
    q += ' ORDER BY mk'
    r = [dict(x) for x in c.execute(q, a)]
    c.close()
    return r


def grid_save(d):
    c = conn()
    c.execute('INSERT INTO grid_reading(kind,meter_no,mk,rdate,val,usage,fee,note,updated) '
              'VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(kind,mk) DO UPDATE SET '
              'meter_no=excluded.meter_no, rdate=excluded.rdate, val=excluded.val, '
              'usage=excluded.usage, fee=excluded.fee, note=excluded.note, updated=excluded.updated',
              (d.get('kind') or 'elec', d.get('meter_no') or '', d['mk'], d.get('rdate') or '',
               d.get('val'), d.get('usage'), d.get('fee'), d.get('note') or '', now()))
    c.commit(); c.close()


def grid_delete(mk, kind='elec'):
    c = conn()
    c.execute('DELETE FROM grid_reading WHERE kind=? AND mk=?', (kind, mk))
    c.commit(); c.close()


def meter_delete(mid):
    c = conn()
    c.execute('DELETE FROM reading WHERE meter_id=?', (mid,))
    c.execute('DELETE FROM meter WHERE id=?', (mid,))
    c.commit(); c.close()


# ================= 抄表记录 =================
def readings(meter_id):
    c = conn()
    r = [dict(x) for x in c.execute(
        'SELECT * FROM reading WHERE meter_id=? ORDER BY rdate', (meter_id,))]
    c.close()
    return r


def reading_save(meter_id, rdate, val, note=''):
    c = conn()
    c.execute('PRAGMA busy_timeout=15000')
    if val is None or val == '':
        c.execute('DELETE FROM reading WHERE meter_id=? AND rdate=?', (meter_id, rdate))
    else:
        c.execute('''INSERT INTO reading(meter_id,rdate,val,note,created) VALUES(?,?,?,?,?)
                     ON CONFLICT(meter_id,rdate) DO UPDATE SET val=excluded.val, note=excluded.note''',
                  (meter_id, rdate, float(val), note, now()))
    c.commit(); c.close()


def all_readings():
    """{(meter_id, rdate): val}"""
    c = conn()
    d = {(r['meter_id'], r['rdate']): r['val'] for r in c.execute('SELECT meter_id,rdate,val FROM reading')}
    c.close()
    return d


# ================= 价格 =================
def prices():
    c = conn()
    d = {r['kind']: r['unit_price'] for r in c.execute('SELECT * FROM price')}
    c.close()
    return d


def price_set(kind, v):
    c = conn()
    c.execute('INSERT INTO price(kind,unit_price,updated) VALUES(?,?,?) '
              'ON CONFLICT(kind) DO UPDATE SET unit_price=excluded.unit_price, updated=excluded.updated',
              (kind, float(v), now()))
    c.commit(); c.close()


def price_month(kind, mk):
    """某月生效单价：优先取按月设置，否则取默认单价"""
    c = conn()
    r = c.execute('SELECT unit_price FROM price_m WHERE kind=? AND mk=?', (kind, mk)).fetchone()
    if r:
        c.close()
        return float(r['unit_price'])
    r = c.execute('SELECT unit_price FROM price WHERE kind=?', (kind,)).fetchone()
    c.close()
    return float(r['unit_price']) if r else 0.0


def price_months(kind):
    """{mk: 单价}"""
    c = conn()
    d = {r['mk']: float(r['unit_price']) for r in c.execute(
        'SELECT mk,unit_price FROM price_m WHERE kind=? ORDER BY mk', (kind,))}
    c.close()
    return d


def price_month_set(kind, mk, v):
    c = conn()
    if v in (None, ''):
        c.execute('DELETE FROM price_m WHERE kind=? AND mk=?', (kind, mk))
    else:
        c.execute('INSERT INTO price_m(kind,mk,unit_price,updated) VALUES(?,?,?,?) '
                  'ON CONFLICT(kind,mk) DO UPDATE SET unit_price=excluded.unit_price, '
                  'updated=excluded.updated', (kind, mk, float(v), now()))
    c.commit(); c.close()


# ================= 计算核心 =================
def _d(s):
    if isinstance(s, datetime.date):
        return s
    return datetime.date(*map(int, str(s)[:10].split('-')))


def meter_usage(meter, rd):
    """
    返回该表的所有"用能区间"列表：
      [{start, end, days, raw, usage, daily, prev_val, curr_val, gap}]
    raw = 读数差；usage = raw × mult
    gap=True 表示跨期过长（>45天）或缺读数，标记为"数据不完整"
    backwards=True 表示读数倒走（换表/抄错），该段用量不计入统计
    reset=True 表示识别为"换表清零"：前一段倒走 + 后一段跳涨，按清零后的表底估算用量
    """
    if len(rd) < 2:
        return []
    pts = [(r['rdate'], r['val']) for r in rd if r['val'] is not None]
    pts.sort()
    mult = float(meter.get('mult') or 1)
    out = []
    for i in range(1, len(pts)):
        pv, cv = pts[i - 1][1], pts[i][1]
        d0, d1 = _d(pts[i - 1][0]), _d(pts[i][0])
        days = (d1 - d0).days
        if days <= 0:
            continue
        raw = cv - pv
        out.append(dict(start=d0.isoformat(), end=d1.isoformat(), days=days,
                        raw=round(raw, 2), usage=round(raw * mult, 2),
                        daily=round(raw * mult / days, 4),
                        prev_val=pv, curr_val=cv,
                        gap=days > 45, backwards=raw < 0, reset=False))

    # 换表清零识别：若某段倒走，且相邻的下一段跳涨（新表从 0 起），
    # 则把下一段也标记为 reset，用量按"新表表底 × 倍率"计（近似，误差仅限换表当次）
    for i, iv in enumerate(out):
        if not iv['backwards']:
            continue
        nxt = out[i + 1] if i + 1 < len(out) else None
        prv = out[i - 1] if i > 0 else None
        # 判断是否为换表：倒走幅度大（>5%）且新表底明显小于原表（<50%）
        if nxt is None or iv['curr_val'] >= iv['prev_val'] * 0.5:
            continue
        if nxt['days'] > 45:
            continue
        # 该段（倒走段）本身不计；下一段按新表表底计
        nxt['reset'] = True
        nxt['reset_note'] = '换表清零：%s 倒走，%s 起新表底 %s' % (
            iv['start'], nxt['start'], nxt['prev_val'])
        # 用同一表在倒走前的日均估算换表期间用量，避免新表底偏小造成漏计
        if prv is not None and prv['days'] > 0 and prv['raw'] > 0:
            est = prv['daily'] * nxt['days']
            nxt['raw'] = round(est / mult, 2)
            nxt['usage'] = round(est, 2)
            nxt['daily'] = round(est / nxt['days'], 4)
    return out


def split_usage(interval, month):
    """
    把一段用能按天均摊到它跨越的各个月份，返回 {month: usage}
    这样周抄、半月抄、跨月抄都能正确归到月度
    """
    d0 = _d(interval['start']); d1 = _d(interval['end'])
    total_days = interval['days']
    if total_days <= 0:
        return {}
    daily = interval['usage'] / total_days
    res = {}
    cur = d0
    while cur < d1:
        nxt = datetime.date(cur.year + (cur.month == 12), (cur.month % 12) + 1, 1)
        seg_end = min(nxt, d1)
        days = (seg_end - cur).days
        if days > 0:
            key = cur.strftime('%Y-%m')
            res[key] = res.get(key, 0) + daily * days
        cur = seg_end
    return res


def build_monthly(meters_list, readings_map, prices_map):
    """
    返回结构化月度数据：
    {meter_id: {'meta':..., 'months': {'2026-05': {'usage','cost','days','cover'}}, 'intervals':[...]}}
    """
    out = {}
    for m in meters_list:
        rd = [{'rdate': k[1], 'val': v} for k, v in readings_map.items() if k[0] == m['id']]
        rd.sort(key=lambda x: x['rdate'])
        ivs = meter_usage(m, rd)
        months = {}
        cover = {}
        for iv in ivs:
            if iv['backwards']:
                continue  # 倒走段（换表/抄错）不计用量，已在下一段按新表底估算
            sp = split_usage(iv, None)
            for mk, uv in sp.items():
                months[mk] = months.get(mk, 0) + uv
                cover[mk] = cover.get(mk, 0) + 1
        unit = prices_map.get(m['kind'], 0)
        for mk in months:
            months[mk] = dict(usage=round(months[mk], 2),
                              cost=round(months[mk] * unit, 2),
                              price=unit)
        out[m['id']] = dict(meta=m, months=months, intervals=ivs)
    return out


def apply_month_price(monthly, kind, mk_price):
    """按当月单价重算金额（mk_price: 已合并了按月价的 {mk: price}）"""
    for v in monthly.values():
        if v['meta']['kind'] != kind:
            continue
        for mk, mv in v['months'].items():
            p = mk_price.get(mk)
            if p is not None and p != mv.get('price'):
                mv['price'] = p
                mv['cost'] = round(mv['usage'] * p, 2)


def month_list(monthly):
    s = set()
    for v in monthly.values():
        s.update(v['months'].keys())
    return sorted(s)


def agg_by(monthly, months, key='sys', kind=None):
    """按 sys/zone 聚合成 {label: {month: usage}}"""
    res = {}
    for mid, v in monthly.items():
        m = v['meta']
        if kind and m['kind'] != kind:
            continue
        if m.get('is_total'):
            continue
        lbl = m.get(key) or '未分类'
        res.setdefault(lbl, {})
        for mk, mv in v['months'].items():
            if mk in months:
                res[lbl][mk] = res[lbl].get(mk, 0) + mv['usage']
    return res


def prev_month(mk):
    y, m = int(mk[:4]), int(mk[5:7])
    if m == 1:
        return '%d-12' % (y - 1)
    return '%d-%02d' % (y, m - 1)


def same_month_last_year(mk):
    return '%d%s' % (int(mk[:4]) - 1, mk[4:])


def days_in_month(mk):
    y, m = int(mk[:4]), int(mk[5:7])
    return calendar.monthrange(y, m)[1]


def log(user, act, detail='', ip=''):
    try:
        c = conn()
        c.execute('INSERT INTO log(ts,user,act,detail,ip) VALUES(?,?,?,?,?)',
                  (now(), user or '', act or '', detail or '', ip or ''))
        c.commit(); c.close()
    except Exception:
        pass
