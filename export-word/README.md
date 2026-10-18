# Xuất Word (.docx) thuần ABAP Cloud — bản thử trên nafoods

Sinh file Word từ **template .docx có `${tag}`** + dữ liệu ABAP, hoàn toàn ở backend (không UI5/BAS).
Fiori Elements tự hiện link tải nhờ `@Semantics.largeObject`.

Chạy thử: nafoods CUS (client 100), package local `$TMP`, mọi object hậu tố `_TEST`.

## Luồng

```
App "Template Word" (ZC_WORD_TMPL_TEST)  ── tải lên .docx, Mã template = PO
App "Xuất Word đơn mua hàng" (ZCE_WORD_PO_TEST, custom entity)
  List Report: danh sách đơn mua (I_PurchaseOrderAPI01) — KHÔNG sinh file
  Mở 1 đơn → UI xin FileContent → ZCL_WORD_PO_QUERY_TEST
      đọc template PO mới nhất + item (I_PurchaseOrderItemAPI01) → ZCL_WORD_DOCX_TEST=>render
      → link "File Word" tải về; lỗi (thiếu template / thiếu tag) hiện ở field "Thông báo"
```

## Object (thứ tự activate)

| # | Object | Vai trò |
|---|---|---|
| 1 | `ZTB_WDTPL_TEST`, `ZTB_WDTPL_TEST_D` | Bảng template + bảng draft |
| 2 | `ZCL_WORD_DOCX_TEST` | **Lõi**: `render( iv_template, is_data ) → ev_docx / ev_error` |
| 3 | `ZR_WORD_TMPL_TEST` → `ZC_WORD_TMPL_TEST` (+ DDLX) | CDS kho template |
| 4 | `ZR_WORD_TMPL_TEST` BDEF → `ZBP_R_WORD_TMPL_TEST` → `ZC_WORD_TMPL_TEST` BDEF | Managed + draft, validation file phải là .docx |
| 5 | `ZCE_WORD_PO_TEST` → `ZCL_WORD_PO_QUERY_TEST` | Custom entity + query provider (activate cùng lúc) |
| 6 | `ZUI_WORD_TMPL_TEST` (SRVD) → tạo tay SRVB `ZUI_WORD_TMPL_TEST_O4` (OData V4 UI) → Publish | Preview 2 entity `Template` và `PurchaseOrder` |

Template mẫu: `template/demo_template.docx` (sinh bằng `python tools/build_template.py template/demo_template.docx`).

## Cú pháp template

| Trong Word | Ý nghĩa |
|---|---|
| `${purchaseOrder}` | Field thường của `is_data`. Tên khớp **bỏ `_`, không phân biệt hoa thường**: `purchase_order` = `${purchaseOrder}` |
| `${table:items.qty}` đặt trong **1 dòng bảng Word** | Nhân dòng đó theo số dòng của component bảng `items`; bảng rỗng → dòng biến mất |
| Tag ở header / footer | Được thay như thân văn bản |

Định dạng giá trị: ngày → `DD.MM.YYYY`; số → `NUMBER = USER`; tiền → **ABAP tự định dạng sẵn thành chuỗi** bằng
`CURRENCY = <loại tiền>` trước khi đưa vào (field `CURR` VND đọc thẳng bằng SQL nhỏ hơn 100 lần — `knowledge/lessons.md`).

Báo lỗi sớm: tag không có dữ liệu → `ev_error = "Thiếu dữ liệu: x, y"`, không sinh file thiếu.

## Giới hạn bản đầu

- Không chèn ảnh động (QR, chữ ký). Logo cố định thì vẽ sẵn trong template.
- Word hay cắt 1 tag thành nhiều đoạn chữ — class tự gộp lại; đoạn gộp lấy định dạng của đoạn chứa `${`.
- Không xử lý bảng lồng trong bảng; tag trong text box chưa thử.
- Custom entity: sort chỉ theo 4 field đẩy xuống SQL (PO, Company Code, NCC, ngày); filter ẩn với field còn lại.

## Trạng thái

- 2026-10-18: **chạy thật trên nafoods** — pull abapGit từ repo `NAF_LOCAL`, activate, tải template, mở đơn
  `4500000004` → file Word đúng (header/footer, bảng item, tổng tiền VND định dạng đúng, tag bị cắt đã gộp).
- Sửa sau lần chạy đầu: mã NCC có khoảng trắng thừa sau `ALPHA = OUT` → bọc `condense( )`.
- Chưa thử: đơn nhiều item, template soạn tay bằng Word (tag bị cắt nhiều kiểu hơn), text box.
