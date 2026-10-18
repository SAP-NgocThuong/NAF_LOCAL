"""Thu template bang ban port Python 1-1 cua ZCL_WORD_DOCX_TEST, khong can he thong SAP.

    python tools/test_render.py template/demo_template.docx template/demo_sign.png out.docx

Muon chac Word mo duoc: mo out.docx bang Word (hoac Word COM SaveAs2 PDF).
"""
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).parent))
import word_port as wp  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

KEYS = dict(is_group=False, is_item=False, is_subtotal=False, group_label='', item='', material='', text='',
            qty='', unit='', price='', amount='')
data = dict(purchase_order='4500000004', company_code='V010', supplier='10000005', supplier_name='Công Ty TNHH Sao Khuê',
            po_date='26.06.2026', currency='VND', total='2.640.000', total_words=wp.words_vn(2640000) + ' đồng',
            has_zero_price=True, note='Giao hàng tại kho công ty V010\nLiên hệ: phòng mua hàng', printed_by='CB9980000015',
            sign=open(sys.argv[2], 'rb').read(),
            items=[dict(KEYS, **it) for it in [
                dict(is_group=True, group_label='Hàng có giá'),
                dict(is_item=True, item='10', material='41000052', text='Test BTP dịch chanh leo VN', qty='30', unit='KG', price='88.000', amount='2.640.000'),
                dict(is_subtotal=True, group_label='Cộng hàng có giá', amount='2.640.000'),
                dict(is_group=True, group_label='Hàng chưa có giá'),
                dict(is_item=True, item='20', material='71000000', text='Vật tư không giá', qty='2', unit='CAI', price='0', amount='0'),
                dict(is_subtotal=True, group_label='Cộng hàng chưa có giá', amount='0')]],
            copies=[dict(label='Liên 1: Lưu', is_last=False), dict(label='Liên 2: Giao nhà cung cấp', is_last=True)])

r = wp.Renderer(open(sys.argv[1], 'rb').read())
out = r.run(data)
if out is None:
    print('Thiếu dữ liệu:', ', '.join(r.missing))
    sys.exit(1)
open(sys.argv[3], 'wb').write(out)
print('da tao', sys.argv[3])
