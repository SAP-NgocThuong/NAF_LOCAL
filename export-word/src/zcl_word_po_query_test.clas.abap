CLASS zcl_word_po_query_test DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.
    INTERFACES if_rap_query_provider.

  PRIVATE SECTION.
    CONSTANTS: c_tmpl_id TYPE ztb_wdtpl_test-tmpl_id VALUE 'PO',
               c_sign_id TYPE ztb_wdtpl_test-tmpl_id VALUE 'SIGN'.

    TYPES tt_result TYPE STANDARD TABLE OF zce_word_po_test WITH EMPTY KEY.

    TYPES: BEGIN OF ty_item,
             is_group    TYPE abap_bool,
             is_item     TYPE abap_bool,
             is_subtotal TYPE abap_bool,
             group_label TYPE string,
             item        TYPE string,
             material    TYPE string,
             text        TYPE string,
             qty         TYPE string,
             unit        TYPE string,
             price       TYPE string,
             amount      TYPE string,
           END OF ty_item,
           tt_item TYPE STANDARD TABLE OF ty_item WITH EMPTY KEY,
           BEGIN OF ty_copy,
             label   TYPE string,
             is_last TYPE abap_bool,
           END OF ty_copy,
           tt_copy TYPE STANDARD TABLE OF ty_copy WITH EMPTY KEY,
           BEGIN OF ty_doc,
             purchase_order TYPE string,
             company_code   TYPE string,
             supplier       TYPE string,
             supplier_name  TYPE string,
             po_date        TYPE d,
             currency       TYPE string,
             total          TYPE string,
             total_words    TYPE string,
             has_zero_price TYPE abap_bool,
             note           TYPE string,
             printed_by     TYPE string,
             sign           TYPE xstring,
             items          TYPE tt_item,
             copies         TYPE tt_copy,
           END OF ty_doc.

    METHODS read_header
      IMPORTING io_request       TYPE REF TO if_rap_query_request
      RETURNING VALUE(rt_result) TYPE tt_result.

    METHODS add_files
      CHANGING ct_result TYPE tt_result.
ENDCLASS.



CLASS zcl_word_po_query_test IMPLEMENTATION.

  METHOD if_rap_query_provider~select.
    DATA(lv_where) = io_request->get_filter( )->get_as_sql_string( ).

    IF io_request->is_total_numb_of_rec_requested( ).
      SELECT COUNT(*) FROM i_purchaseorderapi01
        WHERE (lv_where)
        INTO @DATA(lv_count).
      io_response->set_total_number_of_records( lv_count ).
    ENDIF.

    IF io_request->is_data_requested( ).
      DATA(lt_result) = read_header( io_request ).

      " File chi sinh khi UI xin noi dung file (mo Object Page / bam tai), khong sinh cho ca danh sach
      DATA(lt_requested) = io_request->get_requested_elements( ).
      LOOP AT lt_requested INTO DATA(lv_element).
        IF to_upper( lv_element ) = `FILECONTENT`.
          add_files( CHANGING ct_result = lt_result ).
          EXIT.
        ENDIF.
      ENDLOOP.

      io_response->set_data( lt_result ).
    ENDIF.
  ENDMETHOD.


  METHOD read_header.
    DATA(lv_where) = io_request->get_filter( )->get_as_sql_string( ).

    " Chi sort duoc field co tren I_PurchaseOrderAPI01; PurchaseOrder cuoi cung de thu tu duy nhat khi chia trang
    DATA lt_order TYPE string_table.
    DATA(lt_sort) = io_request->get_sort_elements( ).
    LOOP AT lt_sort INTO DATA(ls_sort)
         WHERE element_name = 'PURCHASEORDER' OR element_name = 'COMPANYCODE'
            OR element_name = 'SUPPLIER' OR element_name = 'PURCHASEORDERDATE'.
      APPEND |{ ls_sort-element_name } { COND #( WHEN ls_sort-descending = abap_true THEN `DESCENDING` ELSE `ASCENDING` ) }| TO lt_order.
    ENDLOOP.
    IF NOT line_exists( lt_sort[ element_name = 'PURCHASEORDER' ] ).
      APPEND `PURCHASEORDER ASCENDING` TO lt_order.
    ENDIF.
    DATA(lv_order) = concat_lines_of( table = lt_order sep = `, ` ).

    DATA(lv_top)  = io_request->get_paging( )->get_page_size( ).
    DATA(lv_skip) = io_request->get_paging( )->get_offset( ).
    IF lv_top = if_rap_query_paging=>page_size_unlimited.
      lv_top = 0.
    ENDIF.

    SELECT PurchaseOrder, CompanyCode, Supplier, PurchaseOrderDate, DocumentCurrency
      FROM i_purchaseorderapi01
      WHERE (lv_where)
      ORDER BY (lv_order)
      INTO CORRESPONDING FIELDS OF TABLE @rt_result
      UP TO @lv_top ROWS OFFSET @lv_skip.

    IF rt_result IS INITIAL.
      RETURN.
    ENDIF.

    SELECT Supplier, SupplierName
      FROM i_supplier
      FOR ALL ENTRIES IN @rt_result
      WHERE Supplier = @rt_result-Supplier
      INTO TABLE @DATA(lt_supplier).

    LOOP AT rt_result ASSIGNING FIELD-SYMBOL(<ls_result>).
      <ls_result>-SupplierName = VALUE #( lt_supplier[ Supplier = <ls_result>-Supplier ]-SupplierName OPTIONAL ).
      <ls_result>-FileMime     = zcl_word_docx_test=>c_mime_docx.
      <ls_result>-FileName     = |PO_{ <ls_result>-PurchaseOrder }.docx|.
    ENDLOOP.
  ENDMETHOD.


  METHOD add_files.
    SELECT tmpl_content
      FROM ztb_wdtpl_test
      WHERE tmpl_id = @c_tmpl_id
      ORDER BY created_at DESCENDING
      INTO @DATA(lv_template)
      UP TO 1 ROWS.
    ENDSELECT.
    IF lv_template IS INITIAL.
      LOOP AT ct_result ASSIGNING FIELD-SYMBOL(<ls_empty>).
        <ls_empty>-Message = |Chưa có template mã { c_tmpl_id } trong app Template Word|.
      ENDLOOP.
      RETURN.
    ENDIF.

    SELECT tmpl_content
      FROM ztb_wdtpl_test
      WHERE tmpl_id = @c_sign_id
      ORDER BY created_at DESCENDING
      INTO @DATA(lv_sign)
      UP TO 1 ROWS.
    ENDSELECT.

    SELECT PurchaseOrder, PurchaseOrderItem, Material, PurchaseOrderItemText,
           OrderQuantity, PurchaseOrderQuantityUnit, NetPriceAmount, NetAmount
      FROM i_purchaseorderitemapi01
      FOR ALL ENTRIES IN @ct_result
      WHERE PurchaseOrder = @ct_result-PurchaseOrder
      INTO TABLE @DATA(lt_item).
    SORT lt_item BY PurchaseOrder PurchaseOrderItem.

    DATA(lv_user) = cl_abap_context_info=>get_user_technical_name( ).

    LOOP AT ct_result ASSIGNING FIELD-SYMBOL(<ls_result>).
      DATA(lv_cur) = <ls_result>-DocumentCurrency.
      DATA(ls_doc) = VALUE ty_doc( purchase_order = <ls_result>-PurchaseOrder
                                   company_code   = <ls_result>-CompanyCode
                                   supplier       = condense( |{ <ls_result>-Supplier ALPHA = OUT }| )
                                   supplier_name  = <ls_result>-SupplierName
                                   po_date        = <ls_result>-PurchaseOrderDate
                                   currency       = lv_cur
                                   note           = |Giao hàng tại kho công ty { <ls_result>-CompanyCode }|
                                                 && |{ cl_abap_char_utilities=>newline }Liên hệ: phòng mua hàng|
                                   printed_by     = lv_user
                                   sign           = lv_sign
                                   copies         = VALUE #( ( label = `Liên 1: Lưu` )
                                                             ( label = `Liên 2: Giao nhà cung cấp` is_last = abap_true ) ) ).

      " Tien CURR doc thang bang SQL: VND luu theo quy uoc 2 so le -> dinh dang qua CURRENCY de dung so le
      " Nhom dong: hang co gia / hang chua co gia, moi nhom co dong tieu de + dong cong nhom
      DATA lv_total TYPE i_purchaseorderitemapi01-NetAmount.
      DATA lv_group_sum TYPE i_purchaseorderitemapi01-NetAmount.
      CLEAR lv_total.
      DO 2 TIMES.
        DATA(lv_priced) = xsdbool( sy-index = 1 ).
        CLEAR lv_group_sum.
        DATA(lt_rows) = VALUE tt_item( ).
        LOOP AT lt_item INTO DATA(ls_item) WHERE PurchaseOrder = <ls_result>-PurchaseOrder.
          IF xsdbool( ls_item-NetPriceAmount <> 0 ) <> lv_priced.
            CONTINUE.
          ENDIF.
          lv_group_sum += ls_item-NetAmount.
          APPEND VALUE #( is_item  = abap_true
                          item     = |{ ls_item-PurchaseOrderItem ALPHA = OUT }|
                          material = |{ ls_item-Material ALPHA = OUT }|
                          text     = ls_item-PurchaseOrderItemText
                          qty      = |{ CONV decfloat34( ls_item-OrderQuantity ) NUMBER = USER }|
                          unit     = ls_item-PurchaseOrderQuantityUnit
                          price    = |{ ls_item-NetPriceAmount CURRENCY = lv_cur NUMBER = USER }|
                          amount   = |{ ls_item-NetAmount CURRENCY = lv_cur NUMBER = USER }| ) TO lt_rows.
        ENDLOOP.
        IF lt_rows IS INITIAL.
          CONTINUE.
        ENDIF.
        IF lv_priced = abap_false.
          ls_doc-has_zero_price = abap_true.
        ENDIF.
        DATA(lv_label) = COND string( WHEN lv_priced = abap_true THEN `Hàng có giá` ELSE `Hàng chưa có giá` ).
        APPEND VALUE #( is_group = abap_true group_label = lv_label ) TO ls_doc-items.
        APPEND LINES OF lt_rows TO ls_doc-items.
        APPEND VALUE #( is_subtotal = abap_true
                        group_label = |Cộng { to_lower( lv_label ) }|
                        amount      = |{ lv_group_sum CURRENCY = lv_cur NUMBER = USER }| ) TO ls_doc-items.
        lv_total += lv_group_sum.
      ENDDO.
      ls_doc-total       = |{ lv_total CURRENCY = lv_cur NUMBER = USER }|.
      ls_doc-total_words = |{ zcl_word_docx_test=>number_to_words_vn( CONV decfloat34( |{ lv_total CURRENCY = lv_cur NUMBER = RAW }| ) ) } đồng|.

      zcl_word_docx_test=>render( EXPORTING iv_template = lv_template
                                            is_data     = ls_doc
                                  IMPORTING ev_docx     = <ls_result>-FileContent
                                            ev_error    = DATA(lv_error) ).
      <ls_result>-Message = lv_error.
    ENDLOOP.
  ENDMETHOD.

ENDCLASS.
