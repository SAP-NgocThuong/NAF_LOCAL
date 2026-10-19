# Thư viện xuất Word (.docx) thuần ABAP Cloud — `ZCL_CORE_WORD_DOCX`

Sinh file Word từ **template .docx có `${tag}`** + dữ liệu ABAP, hoàn toàn ở backend (không UI5/BAS).
Fiori Elements tự hiện link tải nhờ `@Semantics.largeObject`. Repo gốc: `ZCORE_DEV` nhánh `source`
(`word-core/`, `word-demo/`, `word-tools/`); `NAF_LOCAL` chỉ dùng để pull test.

## Cài đặt

| # | Phần | Object | Ghi chú |
|---|---|---|---|
| 1 | **Kho file chung** (của thư viện Excel, `abap-core/`) | `ZXL_FILE`, `ZXL_FILE_D` → `ZR_XL_FILE` → `ZC_XL_FILE` (+DDLX) → BDEF `ZR_XL_FILE` → `ZBP_R_XL_FILE` → BDEF `ZC_XL_FILE` → `ZUI_XL_LIB` + SRVB `ZUI_XL_LIB_O4` | Kho template/ảnh có phiên bản (cùng mã tải lại = version + 1). Nhận xlsx, **docx**, png, jpeg. Tạm dùng chung, sau thiết kế lại |
| 2 | **Lõi Word** (`word-core/`) | `ZCL_CORE_WORD_DOCX` | Cần kho ở bước 1 nếu dùng `render_by_id` / `get_file` |
| 3 | Demo (`word-demo/`, tùy chọn) | `ZCE_CORE_WORD_PO` + `ZCL_CORE_WORD_PO_QUERY` (activate cùng lúc) → `ZUI_CORE_WORD_PO` + SRVB `ZUI_CORE_WORD_PO_O4` | Đơn mua hàng: template mã `PO`, chữ ký mã `SIGN` trong kho file |

## API

```abap
" Template lấy từ kho file theo mã (bản mới nhất)
zcl_core_word_docx=>render_by_id( EXPORTING iv_file_id = 'PO'  is_data = ls_doc
                                  IMPORTING ev_docx = lv_docx  ev_error = lv_error  et_messages = lt_msg ).
" Template tự truyền vào
zcl_core_word_docx=>render( EXPORTING iv_template = lv_xstring  is_data = ls_doc  IMPORTING ... ).
" Ảnh / file bất kỳ trong kho (logo, chữ ký)
ls_doc-sign = zcl_core_word_docx=>get_file( 'SIGN' ).
" Số thành chữ
zcl_core_word_docx=>number_to_words_vn( 2640000 ).   " Hai triệu sáu trăm bốn mươi nghìn
```

## Demo đơn mua hàng

```
Kho file (ZC_XL_FILE): Mã PO = template/demo_template.docx · Mã SIGN = template/demo_sign.png
App ZUI_CORE_WORD_PO_O4 (custom entity ZCE_CORE_WORD_PO, I_PurchaseOrderAPI01 + I_PurchaseOrderItemAPI01)
  Bấm GO → cột Thông báo: sinh thử từng đơn → đỏ lỗi / vàng cảnh báo / xanh Sẵn sàng
  Bấm link File Word → sinh và tải file (2 liên, nhóm dòng, bằng chữ, ghi chú, chữ ký, barcode)
```

Template mẫu: `python tools/build_template.py template/demo_template.docx template/demo_sign.png`
(barcode cần nhúng font — mục Barcode bên dưới).
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

Hàm phụ: `zcl_core_word_docx=>number_to_words_vn( 2640000 )` → `Hai triệu sáu trăm bốn mươi nghìn`.

Định dạng giá trị: ngày → `DD.MM.YYYY`; số → `NUMBER = USER`; tiền → **ABAP tự định dạng sẵn thành chuỗi** bằng
`CURRENCY = <loại tiền>` (field `CURR` VND đọc thẳng bằng SQL nhỏ hơn 100 lần — `knowledge/lessons.md`).
Số thành chữ cho tiền: `|{ amount CURRENCY = cur NUMBER = RAW }|` → `decfloat34` → `number_to_words_vn`.

Kết quả trả về: `ev_docx` (file), `ev_error` (chuỗi lỗi gộp, rỗng = thành công), `et_messages` (bảng `type`/`tag`/`text`:
`E` lỗi → không sinh file · `W` cảnh báo, vẫn sinh (bảng rỗng, ảnh rỗng) · `S` thành công kèm dung lượng).
Hằng `zcl_core_word_docx=>c_msg_type-error|warning|success`.

Báo lỗi sớm — không sinh file khi: tag không có dữ liệu · tên field trong `#if`/`#each`/`rowif` sai · thiếu thẻ đóng ·
ảnh không phải PNG/JPEG. Thông báo: `Thiếu dữ liệu: x, y`.

## Barcode / QR

| Loại | Cách | Trạng thái |
|---|---|---|
| Barcode 1 chiều Code39 | Ô chữ `*${purchaseOrder}*` đặt font **Libre Barcode 39** (Google Fonts, OFL) + **nhúng font vào template** | Đã thử 2026-10-18: gỡ font khỏi máy vẫn hiện barcode (Word dùng font nhúng), xuất PDF đúng. [Unverified] chưa quét thử bằng máy quét |
| QR | Không làm bằng font được — field `DISPLAYBARCODE` của Word hoặc ABAP tự vẽ PNG | Chưa làm |

Nhúng font: cài font → mở template bằng Word → File > Options > Save > **Embed fonts in the file**, bỏ tick
"Embed only the characters used" (lưu đủ ký tự cho mọi số chứng từ) → Save → chạy
`python tools/keep_embedded_font.py vao.docx ra.docx "Libre Barcode 39"` để bỏ Arial… Word nhúng kèm (8 MB → 30 KB).
Code39 chỉ nhận `0-9 A-Z - . $ / + %` và dấu cách; `*` hai đầu là ký tự bắt đầu/kết thúc.

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
- 2026-10-19: **đóng gói thành thư viện** — tiền tố ZCORE (`ZCL_CORE_WORD_DOCX`, demo `ZCE_CORE_WORD_PO` /
  `ZCL_CORE_WORD_PO_QUERY` / `ZUI_CORE_WORD_PO`), bỏ kho riêng `ZTB_WDTPL_TEST` → dùng kho file chung `ZXL_FILE`
  (thêm `render_by_id`, `get_file`). Object `_TEST` cũ: bản sao ở `trash/word-test-objects-2026-10-19/`.
  `[Unverified — chưa activate bản đóng gói]`.
