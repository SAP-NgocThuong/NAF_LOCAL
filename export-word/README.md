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

Template mẫu: `python tools/build_template.py template/demo_template.docx template/demo_sign.png`.
Thử template không cần SAP: `python tools/test_render.py template/demo_template.docx template/demo_sign.png out.docx`
(`tools/word_port.py` = bản port Python 1-1 của class — sửa ABAP thì sửa cả file này).

## Cú pháp template

| Trong Word | Ý nghĩa |
|---|---|
| `${purchaseOrder}` | Field thường. Tên khớp **bỏ `_`, không phân biệt hoa thường**: `purchase_order` = `${purchaseOrder}`. Giá trị có xuống dòng → xuống dòng trong Word |
| `${table:items.qty}` trong **1 dòng bảng Word** | Nhân dòng theo mảng `items`; mảng rỗng → dòng biến mất |
| Nhiều dòng bảng liền nhau cùng `items` + `${rowif:isGroup}` / `${rowifnot:x}` | Mỗi phần tử chọn dòng mẫu hợp với nó → dòng tiêu đề nhóm / dòng hàng / dòng cộng nhóm (như L1/L2 bên Excel) |
| Đoạn `${#each:copies}` … đoạn `${/each:copies}` | Lặp **cả khối** (đoạn văn + bảng) theo mảng; trong khối dùng field của phần tử, không có thì lấy field ngoài. Dùng in nhiều liên / nhiều chứng từ trong 1 file |
| Đoạn `${#if:note}` … `${/if:note}` · `${#ifnot:isLast}` … `${/if:isLast}` | Giữ / bỏ khối theo field (khác rỗng, khác 0; mảng: có dòng) |
| `${image:sign}` · `${image:sign\|3}` | Chèn ảnh PNG/JPEG (field `xstring`), rộng 4 cm mặc định hoặc 3 cm, cao theo tỷ lệ ảnh. Field rỗng → bỏ tag |
| Tag ở header / footer | Thay như thân văn bản (ảnh cũng được) |

Thẻ khối (`#each`, `#if`, `#ifnot`, `/…`) phải nằm **một mình 1 đoạn văn ngoài bảng** — cả đoạn đó bị bỏ khi ghép.

Hàm phụ: `zcl_word_docx_test=>number_to_words_vn( 2640000 )` → `Hai triệu sáu trăm bốn mươi nghìn`.

Định dạng giá trị: ngày → `DD.MM.YYYY`; số → `NUMBER = USER`; tiền → **ABAP tự định dạng sẵn thành chuỗi** bằng
`CURRENCY = <loại tiền>` (field `CURR` VND đọc thẳng bằng SQL nhỏ hơn 100 lần — `knowledge/lessons.md`).
Số thành chữ cho tiền: `|{ amount CURRENCY = cur NUMBER = RAW }|` → `decfloat34` → `number_to_words_vn`.

Kết quả trả về: `ev_docx` (file), `ev_error` (chuỗi lỗi gộp, rỗng = thành công), `et_messages` (bảng `type`/`tag`/`text`:
`E` lỗi → không sinh file · `W` cảnh báo, vẫn sinh (bảng rỗng, ảnh rỗng) · `S` thành công kèm dung lượng).
Hằng `zcl_word_docx_test=>c_msg_type-error|warning|success`.

Báo lỗi sớm — không sinh file khi: tag không có dữ liệu · tên field trong `#if`/`#each`/`rowif` sai · thiếu thẻ đóng ·
ảnh không phải PNG/JPEG. Thông báo: `Thiếu dữ liệu: x, y`.

## Demo đơn mua hàng (template `PO` + ảnh `SIGN`)

Tải trong app Template Word: Mã `PO` = `template/demo_template.docx`, Mã `SIGN` = `template/demo_sign.png` (ảnh chữ ký mẫu).
File ra: 2 liên (mỗi liên 1 trang), bảng nhóm "Hàng có giá / Hàng chưa có giá" + dòng cộng nhóm, tổng + bằng chữ,
cảnh báo đỏ khi có dòng giá 0, ghi chú 2 dòng, ảnh chữ ký.

## Giới hạn

- Không xuất PDF (ABAP Cloud không có công cụ chuyển; cần Adobe Forms).
- Không bảng lồng trong bảng; không lồng 2 khối **cùng tên** (`#if:x` trong `#if:x`); thẻ khối không đặt trong ô bảng.
- Ảnh chỉ ở thân văn bản / header / footer, chưa chèn ảnh theo từng dòng bảng.
- Word hay cắt 1 tag thành nhiều đoạn chữ — class tự gộp; đoạn gộp lấy định dạng của đoạn chứa `${`.
- Text box chưa thử.
- Custom entity: sort chỉ theo 4 field đẩy xuống SQL (PO, Company Code, NCC, ngày); filter ẩn với field còn lại.

## Trạng thái

- 2026-10-18: bản đầu **chạy thật trên nafoods** (pull abapGit `NAF_LOCAL`, đơn `4500000004` ra đúng).
- 2026-10-18: bản mở rộng (xuống dòng, `#each`, `#if`/`#ifnot`, `rowif`, ảnh PNG/JPEG, số thành chữ) — đã kiểm bằng
  **bản port Python 1-1** + **mở bằng Microsoft Word** (2 trang, 2 ảnh, xuất PDF đúng bố cục) và các ca biên
  (bảng rỗng, ảnh rỗng, JPEG, thiếu thẻ đóng, sai tên field, ảnh sai định dạng, 15 ca số thành chữ).
- 2026-10-18: bản mở rộng **chạy thật trên nafoods** — đơn `4500000004` ra 2 liên, nhóm dòng, bằng chữ, ghi chú 2 dòng,
  ảnh chữ ký. Lỗi lần đầu: PCRE ABAP coi `#` là comment → `${#each…}` không nhận → sửa `\#` (`knowledge/lessons.md`).
  Báo lỗi: bấm GO → cột **Thông báo** thử sinh file từng dòng: đỏ = lỗi (không có file) · vàng = cảnh báo · xanh = Sẵn sàng.
  Nội dung file chỉ giữ khi bấm tải. [Inference] GO chậm hơn vì mỗi dòng trong trang sinh thử 1 lần.
