"""Sinh template Word mau (don mua hang) cho ZCL_WORD_DOCX_TEST.

    python tools/build_template.py template/demo_template.docx

Co san 1 tag bi cat thanh nhieu run (${supplier|Name}) de thu buoc gop tag.
"""
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass


def label(p, text, value):
    p.add_run(text).bold = True
    p.add_run(value)


def main(out):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    doc.sections[0].header.paragraphs[0].text = "Company Code: ${companyCode}"
    doc.sections[0].footer.paragraphs[0].text = "In bởi: ${printedBy}"

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("ĐƠN MUA HÀNG")
    r.bold = True
    r.font.size = Pt(16)

    label(doc.add_paragraph(), "Số đơn: ", "${purchaseOrder}")
    label(doc.add_paragraph(), "Ngày: ", "${poDate}")
    p = doc.add_paragraph()
    p.add_run("Nhà cung cấp: ").bold = True
    p.add_run("${supplier} - ${supplier")
    p.add_run("Name}").italic = True
    label(doc.add_paragraph(), "Tiền tệ: ", "${currency}")

    heads = ["Dòng", "Vật tư", "Mô tả", "SL", "ĐVT", "Đơn giá", "Thành tiền"]
    tags = ["item", "material", "text", "qty", "unit", "price", "amount"]
    tbl = doc.add_table(rows=3, cols=len(heads))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(heads):
        c = tbl.rows[0].cells[i]
        c.text = ""
        c.paragraphs[0].add_run(h).bold = True
        tbl.rows[1].cells[i].text = "${table:items.%s}" % tags[i]
    total = tbl.rows[2].cells[0].merge(tbl.rows[2].cells[5])
    total.text = ""
    total.paragraphs[0].add_run("TỔNG CỘNG").bold = True
    tbl.rows[2].cells[6].text = "${total}"

    doc.add_paragraph()
    sign = doc.add_table(rows=1, cols=2)
    for i, s in enumerate(["Người lập", "Người duyệt"]):
        c = sign.rows[0].cells[i]
        c.text = ""
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].add_run(s).bold = True

    doc.save(out)
    print("da tao", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "demo_template.docx")
