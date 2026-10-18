"""Ban port Python 1-1 cua ZCL_WORD_DOCX_TEST (ban mo rong) de thu tren file docx that."""
import io
import re
import struct
import sys
import zipfile
from decimal import Decimal
from xml.sax.saxutils import escape

PAR = re.compile(r'(?s)<w:p(?:\s[^>]*)?(?<!/)>.*?</w:p>')
PAR_OPEN = re.compile(r'<w:p(?:\s[^>]*)?(?<!/)>')
RUN = re.compile(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>')
ROW = re.compile(r'(?s)<w:tr(?:\s[^>]*)?>.*?</w:tr>')
TAG = re.compile(r'\$\{([^}<]*)\}')
BLOCK = re.compile(r'\$\{#(each|if|ifnot):(\w+)\}')
IMG = re.compile(r'\$\{image:(\w+)(?:\|([\d.]+))?\}')
BREAK = '</w:t><w:br/><w:t xml:space="preserve">'
EMU = 360000


def norm(k):
    return k.replace('_', '').upper()


def is_open(s):
    return len(re.findall(r'\$\{', s)) > len(re.findall(r'\$\{[^}]*\}', s)) or s.endswith('$')


def merge_par(p):
    if '$' not in p:
        return p
    runs = list(RUN.finditer(p))
    if len(runs) < 2:
        return p
    txt = [m.group(1) for m in runs]
    ch = [False] * len(txt)
    op = -1
    for i in range(len(txt)):
        if op >= 0:
            txt[op] += txt[i]; txt[i] = ''; ch[op] = ch[i] = True
            if not is_open(txt[op]):
                op = -1
        elif is_open(txt[i]):
            op = i
    for i in range(len(runs) - 1, -1, -1):
        if ch[i]:
            m = runs[i]
            p = p[:m.start()] + '<w:t xml:space="preserve">%s</w:t>' % txt[i] + p[m.end():]
    return p


def merge_split(xml):
    for m in reversed(list(PAR.finditer(xml))):
        xml = xml[:m.start()] + merge_par(m.group(0)) + xml[m.end():]
    return xml


def to_text(v):
    if isinstance(v, bool):
        return 'X' if v else ''
    return str(v)


def find_comp(scope, name):
    for d in scope:
        if isinstance(d, dict):
            for k, v in d.items():
                if norm(k) == norm(name):
                    return (v,)
    return None


def is_true(scope, name):
    r = find_comp(scope, name)
    if r is None:
        return False
    v = r[0]
    if isinstance(v, list):
        return len(v) > 0
    return bool(v)


def scope_values(scope):
    vals = {}
    for d in scope:
        for k, v in d.items():
            if isinstance(v, (list, dict, bytes)):
                continue
            vals.setdefault(norm(k), to_text(v))
    return vals


def par_bounds(xml, pos):
    opens = list(PAR_OPEN.finditer(xml, 0, pos))
    start = opens[-1].start() if opens else pos
    close = xml.find('</w:p>', pos)
    return start, (close + 6 if close >= 0 else len(xml))


class Renderer:
    def __init__(self, tpl):
        self.zin = zipfile.ZipFile(io.BytesIO(tpl))
        self.files = {n: self.zin.read(n) for n in self.zin.namelist()}
        self.missing, self.media, self.ext, self.image = [], [], [], 0

    def miss(self, t):
        if t not in self.missing:
            self.missing.append(t)

    def process(self, xml, scope):
        xml = self.blocks(xml, scope)
        xml = self.rows(xml, scope)
        xml = self.images(xml, scope)
        return self.fill(xml, scope_values(scope), '')

    def blocks(self, xml, scope):
        while True:
            m = BLOCK.search(xml)
            if not m:
                return xml
            kind, name = m.group(1), m.group(2)
            start, inner_start = par_bounds(xml, m.start())
            close = '${/%s:%s}' % ('if' if kind == 'ifnot' else kind, name)
            cpos = xml.find(close, inner_start)
            if cpos < 0:
                self.miss('#%s:%s (thieu %s)' % (kind, name, close))
                xml = xml[:start] + xml[inner_start:]
                continue
            inner_end, end = par_bounds(xml, cpos)
            inner = xml[inner_start:inner_end]
            out = ''
            if kind == 'each':
                r = find_comp(scope, name)
                if r is None or not isinstance(r[0], list):
                    self.miss('#each:' + name)
                else:
                    for line in r[0]:
                        out += self.process(inner, [line] + scope)
            elif find_comp(scope, name) is None:
                self.miss('#%s:%s' % (kind, name)); out = ''
            elif kind == 'if':
                out = inner if is_true(scope, name) else ''
            else:
                out = inner if not is_true(scope, name) else ''
            xml = xml[:start] + out + xml[end:]

    def rows(self, xml, scope):
        rows = []
        for m in ROW.finditer(xml):
            t = re.search(r'\$\{table:(\w+)\.', m.group(0))
            if t:
                rows.append((m.start(), m.end() - m.start(), t.group(1)))
        out_xml = xml
        i = len(rows) - 1
        while i >= 0:
            last = i
            first = i
            while first > 0 and rows[first - 1][2] == rows[first][2] and rows[first - 1][0] + rows[first - 1][1] == rows[first][0]:
                first -= 1
            i = first - 1
            tab = rows[first][2]
            r = find_comp(scope, tab)
            if r is None or not isinstance(r[0], list):
                self.miss('table:' + tab)
                continue
            out = ''
            for line in r[0]:
                ls = [line]
                vals = scope_values(ls)
                for off, ln, _ in rows[first:last + 1]:
                    tpl = xml[off:off + ln]
                    c = re.search(r'\$\{rowif(not)?:(\w+)\}', tpl)
                    if c:
                        if find_comp(ls, c.group(2)) is None:
                            self.miss('rowif:' + c.group(2))
                        tr = is_true(ls, c.group(2))
                        if (not c.group(1) and not tr) or (c.group(1) and tr):
                            continue
                        tpl = re.sub(r'\$\{rowif(not)?:\w+\}', '', tpl)
                    out += self.fill(tpl, vals, 'table:%s.' % tab)
            s = rows[first][0]
            e = rows[last][0] + rows[last][1]
            out_xml = out_xml[:s] + out + out_xml[e:]
        return out_xml

    def images(self, xml, scope):
        for m in reversed(list(IMG.finditer(xml))):
            name = m.group(1)
            w_cm = Decimal(m.group(2)) if m.group(2) else Decimal(4)
            run = ''
            r = find_comp(scope, name)
            if r is None:
                self.miss('image:' + name)
            elif r[0]:
                data = r[0]
                ext, pw, ph = image_size(data)
                if not ext:
                    self.miss('image:%s (khong phai PNG/JPEG)' % name)
                else:
                    self.image += 1
                    iid, rid = 9000 + self.image, 'rIdZimg%d' % self.image
                    cx = int(w_cm * EMU)
                    cy = int(Decimal(cx) * ph / pw) if pw and ph else cx
                    self.media.append((rid, 'media/zimg%d.%s' % (self.image, ext), data))
                    if ext not in self.ext:
                        self.ext.append(ext)
                    run = ('</w:t></w:r><w:r><w:drawing>'
                           '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">'
                           f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{iid}" name="Picture {iid}"/>'
                           '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                           '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                           '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                           f'<pic:nvPicPr><pic:cNvPr id="{iid}" name="zimg{self.image}.{ext}"/><pic:cNvPicPr/></pic:nvPicPr>'
                           f'<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/>'
                           '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                           f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                           '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
                           '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t xml:space="preserve">')
            xml = xml[:m.start()] + run + xml[m.end():]
        return xml

    def fill(self, xml, vals, prefix):
        for m in reversed(list(TAG.finditer(xml))):
            name = m.group(1).strip()
            if prefix:
                if not name.lower().startswith(prefix.lower()):
                    continue
                name = name[len(prefix):]
            elif name.lower().startswith('table:'):
                self.miss(name + ' (ngoai dong bang)'); continue
            v = vals.get(norm(name))
            if v is None:
                self.miss(prefix + name); v = ''
            else:
                v = escape(v).replace('\r\n', '\n').replace('\n', BREAK)
            xml = xml[:m.start()] + v + xml[m.end():]
        return xml

    def flush(self, part):
        if not self.media:
            return
        rels = 'word/_rels/%s.rels' % part[len('word/'):]
        txt = self.files[rels].decode() if rels in self.files else (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        new = ''
        for rid, target, data in self.media:
            new += f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{target}"/>'
            self.files['word/' + target] = data
        self.files[rels] = txt.replace('</Relationships>', new + '</Relationships>').encode()
        self.media = []

    def run(self, data):
        parts = [n for n in self.files if n == 'word/document.xml' or re.match(r'word/(header|footer).*\.xml$', n)]
        for p in parts:
            xml = merge_split(self.files[p].decode('utf-8'))
            self.files[p] = self.process(xml, [data]).encode('utf-8')
            self.flush(p)
        if self.ext:
            ct = self.files['[Content_Types].xml'].decode()
            for e in self.ext:
                if ('extension="%s"' % e) not in ct.lower():
                    ct = ct.replace('</Types>', f'<Default Extension="{e}" ContentType="image/{e}"/></Types>')
            self.files['[Content_Types].xml'] = ct.encode()
        if self.missing:
            return None
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for n, b in self.files.items():
                z.writestr(n, b)
        return out.getvalue()


def image_size(b):
    if len(b) >= 24 and b[:4] == b'\x89PNG':
        return 'png', struct.unpack('>I', b[16:20])[0], struct.unpack('>I', b[20:24])[0]
    if len(b) < 4 or b[:2] != b'\xff\xd8':
        return '', 0, 0
    off = 2
    while off + 9 < len(b):
        if b[off] != 0xFF:
            return 'jpeg', 0, 0
        if b[off + 1] in (0xC0, 0xC1, 0xC2):
            return 'jpeg', struct.unpack('>H', b[off + 7:off + 9])[0], struct.unpack('>H', b[off + 5:off + 7])[0]
        off += 2 + struct.unpack('>H', b[off + 2:off + 4])[0]
    return 'jpeg', 0, 0


DIG = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín']
UNIT = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ', 'triệu tỷ', 'tỷ tỷ']


def triple(n, full):
    h, t, o = n // 100, (n % 100) // 10, n % 10
    w = []
    if h > 0 or full:
        w.append(DIG[h] + ' trăm')
    if t == 0:
        if o > 0 and (h > 0 or full):
            w.append('lẻ')
    elif t == 1:
        w.append('mười')
    else:
        w.append(DIG[t] + ' mươi')
    if o > 0:
        w.append('mốt' if o == 1 and t >= 2 else 'lăm' if o == 5 and t >= 1 else 'tư' if o == 4 and t >= 2 else DIG[o])
    return ' '.join(w)


def words_vn(x):
    v = int(abs(Decimal(x)))
    if v == 0:
        return 'Không'
    g = []
    while v > 0:
        g.append(v % 1000); v //= 1000
    txt = ''
    for i in range(len(g), 0, -1):
        if g[i - 1] > 0:
            txt = ' '.join((txt + ' ' + triple(g[i - 1], i < len(g)) + ' ' + UNIT[i - 1]).split())
    if Decimal(x) < 0:
        txt = 'âm ' + txt
    return txt[0].upper() + txt[1:]
