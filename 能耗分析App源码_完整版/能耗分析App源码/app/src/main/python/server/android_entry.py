# -*- coding: utf-8 -*-
"""安卓端入口：复用原 app.py 全部接口，只替换启动方式与路径"""
import os, threading, json, re, io, csv, time, secrets, datetime, traceback
from http.server import BaseHTTPRequestHandler
from http.server import ThreadingHTTPServer
_WEB = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "assets", "web")
os.environ["ENERGY_WEB_DIR"] = _WEB
os.environ["ENERGY_NO_BROWSER"] = "1"
from .app import H, DB  # 原 app.py，业务代码未改动

class _ReqH(H):
    def address_string(self):
        return "app"

def start(host="127.0.0.1", port=0):
    DB.init()
    srv = ThreadingHTTPServer((host, port), _ReqH)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv.server_address[1]

APP_DIR = os.path.dirname(os.path.abspath(__file__))
_WEB2 = os.path.join(os.path.dirname(APP_DIR), 'web2')
WEB_DIR = _WEB2 if os.path.isdir(_WEB2) else os.path.join(os.path.dirname(APP_DIR), 'web')
VERSION = 'v1.5'
COPYRIGHT = '开心团团'      # 版权署名，固定不可修改
SESS = {}
PUBLIC = {'auth_state', 'auth_setup', 'auth_login', 'ping'}


def rows(c):
    return [dict(r) for r in c]


class H(BaseHTTPRequestHandler):
    server_version = 'EnergyCenter/' + VERSION
    protocol_version = 'HTTP/1.1'

    def log_message(self, fmt, *args):
        pass

    # ---------- 基础 ----------
    def _send(self, code, body, ctype='application/json; charset=utf-8', headers=None):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        tok = getattr(self, '_cookie', None)
        if tok:
            try:
                self.send_header('Set-Cookie', 'ec_tok=%s; Path=/; SameSite=Lax; Max-Age=2592000' % tok)
            except Exception:
                pass
        self.end_headers()
        try:
            self.wfile.write(body)
        except Exception:
            pass

    def ok(self, d=None):
        self._send(200, json.dumps(dict(ok=True, data=d), ensure_ascii=False, default=str))

    def err(self, msg, code=400):
        self._send(code, json.dumps(dict(ok=False, msg=msg), ensure_ascii=False))

    def token(self):
        ck = self.headers.get('Cookie') or ''
        for part in ck.split(';'):
            if '=' in part:
                k, v = part.strip().split('=', 1)
                if k == 'ec_tok':
                    return v
        return None

    # ---------- 请求分发 ----------
    def do_GET(self):
        try:
            u = urlparse(self.path)
            p = u.path
            q = {k: v[0] for k, v in parse_qs(u.query, keep_blank_values=True).items()}
            if p.startswith('/api/'):
                return self.api_get(p[5:], q)
            if p in ('/', '/index.html'):
                return self.static('index.html', 'text/html; charset=utf-8')
            if p == '/favicon.ico':
                return self._send(404, b'')
            if p == '/api':
                return self.ok(dict(version=VERSION))
            if p.startswith('/download/'):
                return self.download(p[10:])
            return self.static(p.lstrip('/'), None)
        except Exception:
            traceback.print_exc()
            self.err('服务器内部错误', 500)

    def do_POST(self):
        try:
            u = urlparse(self.path)
            n = int(self.headers.get('Content-Length') or 0)
            raw = self.rfile.read(n) if n else b''
            if u.path.startswith('/api/upload'):
                return self.upload(raw)
            d = json.loads(raw.decode('utf-8') or '{}')
            return self.api_post(u.path[5:], d)
        except Exception:
            traceback.print_exc()
            self.err('服务器内部错误', 500)

    def static(self, name, ctype):
        p = os.path.join(WEB_DIR, name)
        if not os.path.isfile(p):
            if os.path.isfile(os.path.join(WEB_DIR, 'index.html')):
                p = os.path.join(WEB_DIR, 'index.html')
            else:
                return self._send(404, b'Not Found')
        ext = os.path.splitext(p)[1].lower()
        ct = ctype or {'js': 'application/javascript; charset=utf-8',
                       'css': 'text/css; charset=utf-8',
                       'html': 'text/html; charset=utf-8',
                       'png': 'image/png', 'ico': 'image/x-icon',
                       'svg': 'image/svg+xml', 'webmanifest': 'application/manifest+json; charset=utf-8', 'json': 'application/json'}.get(ext.lstrip('.'), 'application/octet-stream')
        with open(p, 'rb') as f:
            return self._send(200, f.read(), ct)

    def download(self, key):
        item = TEMPFILES.pop(key, None)
        if not item:
            return self._send(404, b'file expired')
        data, fn, ct = item
        self._send(200, data, ct, {
            'Content-Disposition': "attachment; filename*=UTF-8''%s" % quote(fn)})

    def upload(self, raw):
        # multipart 简易解析
        ct = self.headers.get('Content-Type') or ''
        if 'multipart/form-data' not in ct:
            return self.err('请使用文件上传')
        if 'boundary=' not in ct:
            return self.err('上传格式不正确')
        boundary = ct.split('boundary=')[1].split(';')[0].strip().strip('"').encode()
        fname = 'upload.xlsx'
        body = None
        for part in raw.split(b'--' + boundary):
            if b'filename="' not in part:
                continue
            head, sep, data = part.partition(b'\r\n\r\n')
            if not sep:
                continue
            try:
                fn = head.split(b'filename="')[1].split(b'"')[0].decode('utf-8', 'ignore')
            except Exception:
                fn = 'upload.xlsx'
            fname = fn
            # 去掉结尾的 CRLF
            body = data
            while body.endswith(b'\r\n'):
                body = body[:-2]
            break
        if not body:
            return self.err('未读取到文件内容')
        if not body.startswith(b'PK'):
            return self.err('文件格式不对，请上传 .xlsx 文件')
        u = SESS.get(self.token())
        if not u:
            return self.err('未登录', 401)
        try:
            res = IMP.import_workbook(body, "auto")
            total = dict(added=sum(x.get('added', 0) for x in res if x.get('ok')),
                         updated=sum(x.get('updated', 0) for x in res if x.get('ok')),
                         reads=sum(x.get('reads', 0) for x in res if x.get('ok')))
            DB.log(u['username'], '导入Excel', '%s → 新增%d 更新%d 读数%d' %
                   (fname, total['added'], total['updated'], total['reads']), self.client_address[0])
            return self.ok(dict(sheets=res, total=total, file=fname))
        except Exception as e:
            traceback.print_exc()
            return self.err('导入失败：%s' % e)

    # ---------- 鉴权 ----------
    def guard(self, p):
        if p in PUBLIC:
            return None, False
        if not DB.has_admin():
            return None, False
        u = SESS.get(self.token())
        if not u:
            self.err('未登录', 401)
            return None, True
        return u, False

    # ================= GET =================
    def api_get(self, p, q):
        U, stop = self.guard(p)
        if stop:
            return
        if p == 'ping':
            return self.ok(dict(version=VERSION, time=DB.now()))

        if p == 'auth_state':
            return self.ok(dict(need_setup=not DB.has_admin(),
                                logged=bool(SESS.get(self.token())),
                                me=SESS.get(self.token())))

        if p == 'bootstrap':
            a = AN.Analysis()
            return self.ok(dict(
                meters=DB.meters(),
                prices=DB.prices(),
                price_months={k: DB.price_months(k) for k in ('elec', 'water')},
                elec_sys=DB.dict_get('sys_elec'), water_sys=DB.dict_get('sys_water'),
                zones=DB.dict_get('zone'),
                dicts=self._dicts(),
                groups=DB.groups(),
                grid=DB.grid_readings(),
                months=a.months, years=a.years(),
                meta=dict(company=DB.meta_get('company', ''),
                          project=DB.meta_get('project', ''),
                          copyright=DB.meta_get('copyright', COPYRIGHT),
                          title=DB.meta_get('title', '能耗分析系统')),
                today=DB.today(), version=VERSION,
                me=SESS.get(self.token()),
            ))

        if p == 'dicts':
            return self.ok(self._dicts())

        if p == 'groups':
            return self.ok(DB.groups())

        if p == 'meters':
            return self.ok(DB.meters(q.get('kind') or None))

        if p == 'readings':
            return self.ok(DB.readings(int(q['meter_id'])))

        if p == 'reading_matrix':
            mk = q.get('month') or datetime.date.today().strftime('%Y-%m')
            kind = q.get('kind', 'elec')
            return self.ok(dict(month=mk, rows=self._matrix(mk, kind)))

        if p == 'readings_day':
            # 某表某月每天读数（用于日抄明细）
            return self.ok(DB.readings(int(q['meter_id'])))

        if p == 'dashboard':
            a = AN.Analysis()
            kind = q.get('kind', 'elec')
            mk = q.get('month') or (a.months[-1] if a.months else datetime.date.today().strftime('%Y-%m'))
            zones = self._zones(q)
            d = dict(month=mk, months=a.months, kind=kind, zones=zones,
                     cmp=a.compare(kind, mk, zones=zones),
                     struct=a.structure(kind, mk, 'sys', zones=zones),
                     zone=a.structure(kind, mk, 'zone', zones=zones),
                     top=a.top_meters(kind, mk, 8, zones=zones),
                     anomalies=a.anomalies(kind, mk, zones=zones)[:8],
                     savings=a.savings(kind, mk)[:-1],
                     cover=a.coverage(kind, mk, zones=zones),
                     years=a.years(),
                     groups=a.group_series(kind, a.months[-12:], DB.groups()),
                     trend=self._trend(a, kind, zones),
                     sub_total=self._sub_total(a, kind, mk, zones),
                     struct_sum=round(sum(x['usage'] for x in a.structure(kind, mk, 'sys', zones=zones)), 1),
                     cat=a.structure(kind, mk, 'cat', zones=zones))
            if kind == 'water':
                d['reclaimed'] = a.reclaimed(a.months[-12:], zones=zones)
            return self.ok(d)

        if p == 'analysis':
            a = AN.Analysis()
            kind = q.get('kind', 'elec')
            mk = q.get('month') or (a.months[-1] if a.months else '')
            zones = self._zones(q)
            return self.ok(dict(
                month=mk, months=a.months, kind=kind, zones=zones,
                cmp=a.compare(kind, mk, zones=zones),
                struct=a.structure(kind, mk, 'sys', zones=zones),
                zone=a.structure(kind, mk, 'zone', zones=zones),
                top=a.top_meters(kind, mk, 20, zones=zones),
                anomalies=a.anomalies(kind, mk, zones=zones),
                savings=a.savings(kind, mk),
                cover=a.coverage(kind, mk, zones=zones),
                quality=a.quality(),
                groups=a.group_series(kind, a.months, DB.groups()),
                struct_trend=a.struct_trend(kind, zones=zones),
                trend=self._trend(a, kind, zones),
                reclaimed=(a.reclaimed(zones=zones) if kind == 'water' else None),
                sub_total=self._sub_total(a, kind, mk, zones),
                cat=a.structure(kind, mk, 'cat', zones=zones),
            ))

        if p == 'trend':
            a = AN.Analysis()
            return self.ok(self._trend(a, q.get('kind', 'elec'), self._zones(q)))

        if p == 'loss':
            # 损耗分析：电=国网总表 vs 低压总表 vs 分表；水=自来水总表 vs 分表
            a = AN.Analysis()
            kind = q.get('kind', 'elec')
            months = q.get('months')
            ms = [m for m in months.split(',') if m] if months else a.months[-12:]
            return self.ok(dict(months=ms, loss=a.loss(kind, ms, zones=self._zones(q)),
                                reclaimed=(a.reclaimed(ms) if kind == 'water' else None)))

        if p == 'quality':
            return self.ok(AN.Analysis().quality(q.get('kind') or None, q.get('month') or None,
                                                 zone=q.get('zone') or None))

        if p == 'year':
            a = AN.Analysis()
            kind = q.get('kind', 'elec')
            yr = q.get('year') or (a.years()[-1] if a.years() else str(datetime.date.today().year))
            zones = self._zones(q)
            d = a.year_summary(kind, yr, zones=zones)
            # 分系统整年合计（用于年度结构，剔除中水）
            months = ['%04d-%02d' % (int(yr), i) for i in range(1, 13)]
            st = a.struct_trend(kind, months, zones=zones)
            st = [dict(label=x['label'], usage=x['total']) for x in st if x['total']]
            tot = sum(x['usage'] for x in st) or 1
            for x in st:
                x['share'] = round(x['usage'] / tot * 100, 1)
            st.sort(key=lambda x: -x['usage'])
            d['struct'] = st
            d['cmp'] = self._year_cmp(a, kind, yr, zones)
            if kind == 'water':
                d['reclaimed'] = a.reclaimed(months, zones=zones)
            return self.ok(d)

        if p == 'groups_trend':
            a = AN.Analysis()
            kind = q.get('kind', 'elec')
            months = q.get('months')
            ms = [m for m in months.split(',') if m] if months else a.months
            return self.ok(dict(months=ms, groups=a.group_series(kind, ms, DB.groups())))

        if p == 'grid':
            return self.ok(DB.grid_readings(q.get('kind') or None))

        if p == 'logs':
            c = DB.conn()
            r = [dict(x) for x in c.execute('SELECT * FROM log ORDER BY id DESC LIMIT 200')]
            c.close()
            return self.ok(r)

        if p == 'export_report':
            a = AN.Analysis()
            kind = q.get('kind', 'elec')
            mk = q.get('month') or a.months[-1]
            year = q.get('year') or mk[:4]
            data = RPT.build(a, kind, mk, year=year)
            k = self._tmp(data, '能耗分析报表_%s_%s.xlsx' % ('电' if kind == 'elec' else '水', mk),
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            return self.ok(dict(url='/download/' + k))

        if p == 'export_fill_template':
            data = IMP.build_fill_template(month=q.get('month') or None,
                                           days=int(q.get('days') or 0) or None)
            k = self._tmp(data, '抄表填空模板（只填数字）.xlsx',
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            return self.ok(dict(url='/download/' + k))

        if p == 'export_template':
            data = IMP.build_template()
            k = self._tmp(data, '能耗抄表导入模板.xlsx',
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            return self.ok(dict(url='/download/' + k))

        if p == 'export_data':
            data = RPT.export_data()
            k = self._tmp(data, '能耗数据明细.xlsx',
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            return self.ok(dict(url='/download/' + k))

        if p == 'export_ledger':
            data = IMP.export_ledger()
            k = self._tmp(data, '表计台账.xlsx',
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            return self.ok(dict(url='/download/' + k))

        if p == 'export_readings':
            data = IMP.export_readings()
            k = self._tmp(data, '抄表读数明细.xlsx',
                          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            return self.ok(dict(url='/download/' + k))

        if p == 'backup':
            data = RPT.backup_json()
            k = self._tmp(data, '能耗数据备份_%s.json' % datetime.date.today().isoformat(),
                          'application/json')
            return self.ok(dict(url='/download/' + k))

        return self.err('未知接口: ' + p, 404)

    def _sub_total(self, a, kind, mk, zones):
        """分表合计 vs 总表（用于总览“是否重复计量”自查）"""
        t = a.total(kind, [mk], zones=zones).get(mk, {})
        s = a.total(kind, [mk], zones=zones, basis='sub').get(mk, {})
        tot, sub = t.get('usage', 0), s.get('usage', 0)
        return dict(month=mk, total=round(tot, 1), sub=round(sub, 1),
                    diff=round(tot - sub, 1),
                    gap=round((tot - sub) / sub * 100, 1) if sub else None)

    def _dicts(self):
        """分类字典（带 id，供界面增删改）"""
        c = DB.conn()
        r = [dict(x) for x in c.execute('SELECT * FROM dict ORDER BY kind, sort, id')]
        c.close()
        return r

    def _zones(self, q):
        """从查询参数解析区域多选（逗号分隔），返回 None 表示不过滤"""
        z = (q.get('zones') or '').strip()
        if not z:
            return None
        return [x for x in z.split(',') if x]

    def _year_cmp(self, a, kind, yr, zones=None):
        """年度对比：本年度合计 vs 上一年度合计（仅比有数据的月份，保证可比）"""
        tt = a.total(kind, zones=zones)
        cur_u = prev_u = 0.0
        cur_c = prev_c = 0.0
        nm = 0
        for mo in range(1, 13):
            mk = '%04d-%02d' % (int(yr), mo)
            v = tt.get(mk)
            if not v:
                continue
            nm += 1
            cur_u += v['usage']; cur_c += v['cost']
            p = tt.get('%04d-%02d' % (int(yr) - 1, mo))
            if p:
                prev_u += p['usage']; prev_c += p['cost']
        # 若去年无任何数据，同比为空
        last_has = any(tt.get('%04d-%02d' % (int(yr) - 1, mo)) for mo in range(1, 13))
        return dict(cur=round(cur_u, 1), cost=round(cur_c, 1), months=nm,
                    prev=round(prev_u, 1) if last_has else None,
                    prev_cost=round(prev_c, 1) if last_has else None,
                    mom=(round((cur_u - prev_u) / prev_u * 100, 1) if last_has and prev_u else None),
                    last_year=int(yr) - 1 if last_has else None)

    def _matrix(self, mk, kind, day=''):
        """抄表矩阵：每表带上次读数、本月末读数、以及本月内每天的读数（供日抄/多次抄）
        day 为空 → 返回整月所有抄表点 days=[{rdate,val,usage}]"""
        c = DB.conn()
        ms = [dict(r) for r in c.execute(
            'SELECT * FROM meter WHERE kind=? AND enabled=1 ORDER BY sys,zone,sort,id', (kind,))]
        out = []
        y, mo = int(mk[:4]), int(mk[5:7])
        ndays = DB.days_in_month(mk)
        month_end = '%s-%02d' % (mk, ndays)
        for m in ms:
            # 上月月底（含）之前最后一次读数
            prev = c.execute('SELECT rdate,val FROM reading WHERE meter_id=? AND rdate<? '
                             'ORDER BY rdate DESC LIMIT 1', (m['id'], mk + '-01')).fetchone()
            # 本月所有读数点
            pts = [dict(rdate=r['rdate'], val=r['val']) for r in c.execute(
                'SELECT rdate,val FROM reading WHERE meter_id=? AND rdate>=? AND rdate<=? '
                'ORDER BY rdate', (m['id'], mk + '-01', month_end))]
            # 按时间顺序算每次用量（含上月基准）
            base = dict(rdate=prev['rdate'], val=prev['val']) if prev else None
            mult = float(m['mult'] or 1)
            last = base
            for pt in pts:
                if last and pt['val'] is not None and last['val'] is not None:
                    pt['usage'] = round((pt['val'] - last['val']) * mult, 2)
                    pt['days'] = (datetime.date(*map(int, pt['rdate'].split('-'))) -
                                  datetime.date(*map(int, last['rdate'].split('-')))).days
                else:
                    pt['usage'] = None
                    pt['days'] = None
                last = pt
            cur = pts[-1] if pts else None
            out.append(dict(id=m['id'], name=m['name'], code=m['code'], sys=m['sys'],
                            zone=m['zone'], mult=m['mult'], total=bool(m.get('is_total')),
                            prev_val=prev['val'] if prev else None,
                            prev_date=prev['rdate'] if prev else None,
                            val=cur['val'] if cur else None,
                            rdate=cur['rdate'] if cur else None,
                            days=pts, note=''))
        c.close()
        return out

    def _trend(self, a, kind, zones=None):
        tt = a.total(kind, zones=zones)
        out = []
        for mk in a.months:
            v = tt.get(mk, {})
            out.append(dict(month=mk, usage=v.get('usage', 0), cost=v.get('cost', 0),
                            daily=round(v.get('usage', 0) / DB.days_in_month(mk), 2)))
        return out

    def _tmp(self, data, fn, ct):
        k = secrets.token_hex(8)
        TEMPFILES[k] = (data, fn, ct)
        if len(TEMPFILES) > 40:
            for kk in list(TEMPFILES)[:10]:
                TEMPFILES.pop(kk, None)
        return k

    # ================= POST =================
    def api_post(self, p, d):
        if p == 'auth_setup':
            if DB.has_admin():
                return self.err('已设置密码')
            pwd = d.get('password') or ''
            if len(pwd) < 4:
                return self.err('密码至少 4 位')
            DB.set_pwd(pwd)
            tok = self._login(d.get('username') or 'admin')
            return self.ok(dict(token=tok))

        if p == 'auth_login':
            if not DB.verify_pwd(d.get('password') or ''):
                return self.err('密码不正确', 401)
            tok = self._login('admin')
            return self.ok(dict(token=tok))

        if p == 'auth_logout':
            SESS.pop(self.token(), None)
            return self.ok()

        U, stop = self.guard(p)
        if stop:
            return

        if p == 'auth_change':
            if not DB.verify_pwd(d.get('old') or ''):
                return self.err('原密码不正确')
            np = d.get('new') or ''
            if len(np) < 4:
                return self.err('新密码至少 4 位')
            DB.set_pwd(np)
            SESS.clear()
            DB.log('admin', '修改密码', '', self.client_address[0])
            return self.ok()

        if p == 'meter':
            mid = DB.meter_save(d)
            DB.log(U['username'], '保存表计', d.get('name', ''), self.client_address[0])
            return self.ok(dict(id=mid))

        if p == 'meter_delete':
            DB.meter_delete(int(d['id']))
            DB.log(U['username'], '删除表计', str(d.get('name', '')), self.client_address[0])
            return self.ok()

        if p == 'reading':
            DB.reading_save(int(d['meter_id']), d['rdate'], d.get('val'), d.get('note', ''))
            return self.ok()

        if p == 'readings_batch':
            n = 0
            for r in d.get('rows', []):
                DB.reading_save(int(r['meter_id']), r['rdate'], r.get('val'), r.get('note', ''))
                n += 1
            DB.log(U['username'], '批量抄表', '%d 条' % n, self.client_address[0])
            return self.ok(dict(n=n))

        if p == 'grid_save':
            DB.grid_save(d)
            DB.log(U['username'], '录入国网抄表', '%s %s' % (d.get('kind', 'elec'), d.get('mk')),
                   self.client_address[0])
            return self.ok()

        if p == 'grid_delete':
            DB.grid_delete(d.get('mk'), d.get('kind') or 'elec')
            return self.ok()

        if p == 'price':
            for k, v in (d or {}).items():
                DB.price_set(k, v)
            return self.ok()

        if p == 'price_month':
            # {kind, mk, value}  或 {kind, items: {mk: value}}；value 为空表示恢复默认
            kind = d.get('kind') or 'elec'
            if 'items' in d:
                for mk, v in (d['items'] or {}).items():
                    DB.price_month_set(kind, mk, v)
            else:
                DB.price_month_set(kind, d.get('mk'), d.get('value'))
            DB.log(U['username'], '设置月度单价', '%s %s' % (kind, d.get('mk') or '批量'),
                   self.client_address[0])
            return self.ok()

        if p == 'import_xlsx':
            import base64 as _b64
            import importer
            data = _b64.b64decode(d.get('content') or '')
            res = IMP.import_workbook(data, "auto")
            total = dict(added=sum(x.get('added', 0) for x in res if x.get('ok')),
                         updated=sum(x.get('updated', 0) for x in res if x.get('ok')),
                         reads=sum(x.get('reads', 0) for x in res if x.get('ok')))
            DB.log(u['username'], '导入Excel(手机)', '%s → 新增%d 更新%d 读数%d' %
                   (d.get('filename') or '', total['added'], total['updated'], total['reads']),
                   self.client_address[0])
            return self.ok(dict(sheets=res, total=total))
        if p == 'import_readings':
            import base64 as _b64
            data = _b64.b64decode(d.get('content') or '')
            r, m = IMP.import_readings(data)
            DB.log(u['username'], '导入读数明细(手机)', '读数%d 未匹配%d' % (r, m), self.client_address[0])
            return self.ok(dict(reads=r, missed=m))
        if p == 'dict_add':
            ok = DB.dict_add(d.get('kind'), d.get('name'))
            DB.log(U['username'], '新增分类', '%s / %s' % (d.get('kind'), d.get('name', '')),
                   self.client_address[0])
            return self.ok(dict(ok=ok)) if ok else self.err('名称不能为空')

        if p == 'dict_del':
            ok, msg = DB.dict_del(int(d['id']))
            if not ok:
                return self.err(msg)
            DB.log(U['username'], '删除分类', str(d.get('name', '')), self.client_address[0])
            return self.ok()

        if p == 'dict_rename':
            ok, msg = DB.dict_rename(int(d['id']), d.get('name'))
            if not ok:
                return self.err(msg)
            DB.log(U['username'], '修改分类', str(d.get('name', '')), self.client_address[0])
            return self.ok()

        if p == 'group':
            gid = DB.group_save(d)
            DB.log(U['username'], '保存显示分组', d.get('name', ''), self.client_address[0])
            return self.ok(dict(id=gid))

        if p == 'group_delete':
            DB.group_delete(int(d['id']))
            return self.ok()

        if p == 'settings':
            # 版权固定为「开心团团」，不允许修改
            for k in ('company', 'project', 'title'):
                if k in d:
                    DB.meta_set(k, d[k])
            DB.meta_set('copyright', COPYRIGHT)
            DB.log(U['username'], '修改设置', '', self.client_address[0])
            return self.ok()

        if p == 'restore':
            try:
                RPT.restore_json(d.get('content') or '')
                return self.ok()
            except Exception as e:
                return self.err('备份文件无法识别：%s' % e)

        if p == 'import_ledger':
            import base64
            try:
                data = base64.b64decode(d.get('content') or '')
                a, u2, s = IMP.import_ledger(data)
                DB.log(U['username'], '导入台账', '新增%d 更新%d' % (a, u2), self.client_address[0])
                return self.ok(dict(added=a, updated=u2))
            except Exception as e:
                traceback.print_exc()
                return self.err('台账导入失败：%s' % e)

        if p == 'import_readings':
            import base64
            try:
                data = base64.b64decode(d.get('content') or '')
                if d.get('mode') == 'check':
                    res = IMP.precheck_readings(data)
                    DB.log(U['username'], '读数预检', '共 %d 行：正常 %d / 待确认 %d / 异常 %d' % (
                        res['stats']['total'], res['stats']['ok'], res['stats']['warn'], res['stats']['err']),
                        self.client_address[0])
                    return self.ok(res)
                n, skipped, kept, miss, same = IMP.import_readings_checked(
                    data, d.get('skip') or [], overwrite=bool(d.get('overwrite', True)))
                DB.log(U['username'], '导入读数',
                       '写入 %d（覆盖模式=%s），跳过 %d，保留原值 %d，未导入 %d，无需变更 %d' % (
                           n, '是' if d.get('overwrite', True) else '否', skipped, kept, miss, same),
                       self.client_address[0])
                return self.ok(dict(reads=n, skipped=skipped, kept=kept,
                                    missed=miss, same=same))
            except Exception as e:
                traceback.print_exc()
                return self.err('读数导入失败：%s' % e)

        if p == 'log':
            DB.log(U['username'], d.get('act', ''), d.get('detail', ''), self.client_address[0])
            return self.ok()

        return self.err('未知接口: ' + p, 404)

    def _login(self, uname):
        tok = secrets.token_hex(16)
        SESS[tok] = dict(username=uname, ts=time.time())
        self._cookie = tok
        return tok


TEMPFILES = {}


def pick_port(start=8830):
    import socket
    for p in list(range(start, start + 40)) + [0]:
        s = socket.socket()
        try:
            s.bind(('127.0.0.1', p))
            port = s.getsockname()[1]
            s.close()
            return port
        except OSError:
            try:
                s.close()
            except Exception:
                pass
    return 0


def main():
    DB.init()
    base = int(os.environ.get('ENERGY_PORT') or 8830)
    srv = None
    for _ in range(3):
        p = pick_port(base)
        try:
            srv = ThreadingHTTPServer(('0.0.0.0', p), H)
            break
        except OSError:
            base += 20
    if srv is None:
        print('  启动失败：端口被占用')
        input('  按回车键退出...')
        return
    real = srv.server_address[1]
    _ips = []
    try:
        import socket as _so
        _s = _so.socket(_so.AF_INET, _so.SOCK_DGRAM); _s.connect(('223.5.5.5', 80))
        _ips.append(_s.getsockname()[0]); _s.close()
    except Exception:
        pass
    try:
        import subprocess as _sp
        _o = _sp.run(['ipconfig'], capture_output=True, text=True, timeout=5).stdout
        _ips += re.findall(r'IPv4[^:]*:\s*((?:192\.168|10\.|172\.)\d+\.\d+)', _o)
    except Exception:
        pass
    print('=' * 48)
    print('   能耗分析系统（手机版）  已启动')
    print('   电脑访问：http://127.0.0.1:%d' % real)
    for _ip in dict.fromkeys(_ips):
        print('   ★ 手机访问：http://%s:%d （需与电脑同一 Wi-Fi）' % (_ip, real))
    print('   数据保存在本程序旁边的 data\\energy.db')
    print('   ★ 关闭本窗口即退出系统（数据已自动保存）')
    print('=' * 48)
    if os.environ.get('ENERGY_NO_BROWSER') != '1':
        threading.Timer(1.2, lambda: webbrowser.open('http://127.0.0.1:%d/' % real)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    except Exception:
        traceback.print_exc()
        input('  程序异常，按回车键退出...')


if __name__ == '__main__':
    main()
