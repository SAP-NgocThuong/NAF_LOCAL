"""Giu lai font nhung cua 1 font (theo ten), bo cac font nhung khac trong docx.

    python tools/keep_embedded_font.py vao.docx ra.docx "Libre Barcode 39"

Word nhung ca Arial... (~8 MB) khi bat Embed fonts; script nay chi giu font barcode (~15 KB).
"""
import io, re, sys, zipfile
src, dst, keep = sys.argv[1], sys.argv[2], sys.argv[3]
zi = zipfile.ZipFile(src)
files = {n: zi.read(n) for n in zi.namelist()}
ft = files['word/fontTable.xml'].decode('utf-8')
rels = files['word/_rels/fontTable.xml.rels'].decode('utf-8')
keep_ids = set()
def fix_font(m):
    block = m.group(0)
    name = re.search(r'w:name="([^"]*)"', block).group(1)
    if name == keep:
        keep_ids.update(re.findall(r'r:id="([^"]+)"', block))
        return block
    return re.sub(r'<w:embed\w+ [^>]*/>', '', block)
ft = re.sub(r'(?s)<w:font [^>]*>.*?</w:font>', fix_font, ft)
drop = []
def fix_rel(m):
    rid, target = m.group(1), m.group(2)
    if rid in keep_ids:
        return m.group(0)
    drop.append('word/' + target)
    return ''
rels = re.sub(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*/>', fix_rel, rels)
files['word/fontTable.xml'] = ft.encode('utf-8')
files['word/_rels/fontTable.xml.rels'] = rels.encode('utf-8')
for d in drop:
    files.pop(d, None)
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
    for n, b in files.items():
        zo.writestr(n, b)
print('giu', keep_ids, 'bo', drop)
