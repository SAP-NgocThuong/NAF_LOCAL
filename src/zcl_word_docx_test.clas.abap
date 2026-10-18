CLASS zcl_word_docx_test DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.
    CONSTANTS c_mime_docx TYPE string VALUE `application/vnd.openxmlformats-officedocument.wordprocessingml.document`.

    CLASS-METHODS render
      IMPORTING iv_template TYPE xstring
                is_data     TYPE any
      EXPORTING ev_docx     TYPE xstring
                ev_error    TYPE string.

    CLASS-METHODS is_docx
      IMPORTING iv_content   TYPE xstring
      RETURNING VALUE(rv_ok) TYPE abap_bool.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_value,
             key   TYPE string,
             value TYPE string,
           END OF ty_value,
           tt_value TYPE HASHED TABLE OF ty_value WITH UNIQUE KEY key.

    CLASS-METHODS merge_split_tags
      IMPORTING iv_xml        TYPE string
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS merge_paragraph
      IMPORTING iv_par        TYPE string
      RETURNING VALUE(rv_par) TYPE string.

    CLASS-METHODS is_open
      IMPORTING iv_text        TYPE string
      RETURNING VALUE(rv_open) TYPE abap_bool.

    CLASS-METHODS expand_rows
      IMPORTING iv_xml        TYPE string
                is_data       TYPE any
      CHANGING  ct_missing    TYPE string_table
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS fill_tags
      IMPORTING iv_xml        TYPE string
                it_values     TYPE tt_value
                iv_prefix     TYPE string OPTIONAL
      CHANGING  ct_missing    TYPE string_table
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS to_values
      IMPORTING is_data          TYPE any
      RETURNING VALUE(rt_values) TYPE tt_value.

    CLASS-METHODS to_text
      IMPORTING ia_value       TYPE any
      RETURNING VALUE(rv_text) TYPE string.

    CLASS-METHODS norm_key
      IMPORTING iv_name       TYPE string
      RETURNING VALUE(rv_key) TYPE string.

    CLASS-METHODS add_missing
      IMPORTING iv_tag     TYPE string
      CHANGING  ct_missing TYPE string_table.
ENDCLASS.



CLASS zcl_word_docx_test IMPLEMENTATION.

  METHOD render.
    " Thay ${tag} trong document.xml + header/footer, nhan dong bang co ${table:<bang>.<field>}
    CLEAR: ev_docx, ev_error.
    IF is_docx( iv_template ) = abap_false.
      ev_error = `File template không phải .docx`.
      RETURN.
    ENDIF.

    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( EXPORTING  zip             = iv_template
                  EXCEPTIONS zip_parse_error = 1
                             OTHERS          = 2 ).

    DATA lt_parts TYPE string_table.
    LOOP AT lo_zip->files INTO DATA(ls_file)
         WHERE name = `word/document.xml` OR name CP 'word/header*.xml' OR name CP 'word/footer*.xml'.
      APPEND ls_file-name TO lt_parts.
    ENDLOOP.

    DATA(lt_values) = to_values( is_data ).
    DATA lt_missing TYPE string_table.

    LOOP AT lt_parts INTO DATA(lv_part).
      DATA lv_raw TYPE xstring.
      lo_zip->get( EXPORTING  name                    = lv_part
                   IMPORTING  content                 = lv_raw
                   EXCEPTIONS zip_index_error         = 1
                              zip_decompression_error = 2
                              OTHERS                  = 3 ).
      IF sy-subrc <> 0.
        ev_error = |Không đọc được { lv_part }|.
        RETURN.
      ENDIF.

      TRY.
          DATA(lv_xml) = cl_abap_conv_codepage=>create_in( )->convert( lv_raw ).
          lv_xml = merge_split_tags( lv_xml ).
          lv_xml = expand_rows( EXPORTING iv_xml     = lv_xml
                                          is_data    = is_data
                                CHANGING  ct_missing = lt_missing ).
          lv_xml = fill_tags( EXPORTING iv_xml     = lv_xml
                                        it_values  = lt_values
                              CHANGING  ct_missing = lt_missing ).
          lv_raw = cl_abap_conv_codepage=>create_out( )->convert( lv_xml ).
        CATCH cx_sy_conversion_codepage.
          ev_error = |Lỗi mã hóa ký tự ở { lv_part }|.
          RETURN.
      ENDTRY.

      lo_zip->delete( EXPORTING  name            = lv_part
                      EXCEPTIONS zip_index_error = 1
                                 OTHERS          = 2 ).
      lo_zip->add( name    = lv_part
                   content = lv_raw ).
    ENDLOOP.

    IF lt_missing IS NOT INITIAL.
      ev_error = |Thiếu dữ liệu: { concat_lines_of( table = lt_missing sep = `, ` ) }|.
      RETURN.
    ENDIF.

    ev_docx = lo_zip->save( ).
  ENDMETHOD.


  METHOD is_docx.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( EXPORTING  zip             = iv_content
                  EXCEPTIONS zip_parse_error = 1
                             OTHERS          = 2 ).
    rv_ok = xsdbool( sy-subrc = 0 AND line_exists( lo_zip->files[ name = `word/document.xml` ] ) ).
  ENDMETHOD.


  METHOD merge_split_tags.
    " Word hay cat 1 tag thanh nhieu <w:t> (soat chinh ta, doi dinh dang) - gop lai theo tung doan van
    rv_xml = iv_xml.
    FIND ALL OCCURRENCES OF PCRE `(?s)<w:p(?:\s[^>]*)?(?<!/)>.*?</w:p>` IN iv_xml RESULTS DATA(lt_par).

    DATA(lv_i) = lines( lt_par ).
    WHILE lv_i > 0.
      DATA(ls_par) = lt_par[ lv_i ].
      lv_i -= 1.
      DATA(lv_old) = substring( val = rv_xml off = ls_par-offset len = ls_par-length ).
      DATA(lv_new) = merge_paragraph( lv_old ).
      IF lv_new <> lv_old.
        rv_xml = replace( val = rv_xml off = ls_par-offset len = ls_par-length with = lv_new ).
      ENDIF.
    ENDWHILE.
  ENDMETHOD.


  METHOD merge_paragraph.
    rv_par = iv_par.
    IF find( val = iv_par sub = `$` ) < 0.
      RETURN.
    ENDIF.

    FIND ALL OCCURRENCES OF PCRE `<w:t(?:\s[^>]*)?>([^<]*)</w:t>` IN iv_par RESULTS DATA(lt_run).
    IF lines( lt_run ) < 2.
      RETURN.
    ENDIF.

    DATA lt_text TYPE string_table.
    LOOP AT lt_run INTO DATA(ls_run).
      DATA(ls_sub) = ls_run-submatches[ 1 ].
      APPEND substring( val = iv_par off = ls_sub-offset len = ls_sub-length ) TO lt_text.
    ENDLOOP.

    " Run mo tag chua dong -> keo text cac run sau vao cho toi khi dong
    DATA lt_changed TYPE STANDARD TABLE OF abap_bool WITH EMPTY KEY.
    lt_changed = VALUE #( FOR i = 1 UNTIL i > lines( lt_text ) ( abap_false ) ).
    DATA(lv_open) = 0.
    LOOP AT lt_text ASSIGNING FIELD-SYMBOL(<lv_text>).
      DATA(lv_idx) = sy-tabix.
      IF lv_open > 0.
        lt_text[ lv_open ] = lt_text[ lv_open ] && <lv_text>.
        CLEAR <lv_text>.
        lt_changed[ lv_open ] = abap_true.
        lt_changed[ lv_idx ]  = abap_true.
        IF is_open( lt_text[ lv_open ] ) = abap_false.
          lv_open = 0.
        ENDIF.
      ELSEIF is_open( <lv_text> ) = abap_true.
        lv_open = lv_idx.
      ENDIF.
    ENDLOOP.

    DATA(lv_i) = lines( lt_run ).
    WHILE lv_i > 0.
      IF lt_changed[ lv_i ] = abap_true.
        ls_run = lt_run[ lv_i ].
        rv_par = replace( val  = rv_par
                          off  = ls_run-offset
                          len  = ls_run-length
                          with = |<w:t xml:space="preserve">{ lt_text[ lv_i ] }</w:t>| ).
      ENDIF.
      lv_i -= 1.
    ENDWHILE.
  ENDMETHOD.


  METHOD is_open.
    rv_open = xsdbool( count( val = iv_text pcre = `\$\{` ) > count( val = iv_text pcre = `\$\{[^}]*\}` )
                       OR matches( val = iv_text pcre = `.*\$` ) ).
  ENDMETHOD.


  METHOD expand_rows.
    " Dong bang Word chua ${table:items.x} -> nhan theo so dong cua component bang 'items'; bang rong -> bo dong
    rv_xml = iv_xml.
    FIND ALL OCCURRENCES OF PCRE `(?s)<w:tr(?:\s[^>]*)?>.*?</w:tr>` IN iv_xml RESULTS DATA(lt_row).
    DATA(lo_type) = cl_abap_typedescr=>describe_by_data( is_data ).
    IF lt_row IS INITIAL OR lo_type->kind <> cl_abap_typedescr=>kind_struct.
      RETURN.
    ENDIF.

    DATA(lo_struct) = CAST cl_abap_structdescr( lo_type ).
    FIELD-SYMBOLS <lt_lines> TYPE ANY TABLE.

    DATA(lv_i) = lines( lt_row ).
    WHILE lv_i > 0.
      DATA(ls_row) = lt_row[ lv_i ].
      lv_i -= 1.
      DATA(lv_row) = substring( val = rv_xml off = ls_row-offset len = ls_row-length ).

      FIND FIRST OCCURRENCE OF PCRE `\$\{table:(\w+)\.` IN lv_row SUBMATCHES DATA(lv_table).
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      UNASSIGN <lt_lines>.
      LOOP AT lo_struct->components INTO DATA(ls_comp) WHERE type_kind = cl_abap_typedescr=>typekind_table.
        IF norm_key( CONV #( ls_comp-name ) ) = norm_key( lv_table ).
          ASSIGN COMPONENT ls_comp-name OF STRUCTURE is_data TO <lt_lines>.
          EXIT.
        ENDIF.
      ENDLOOP.
      IF <lt_lines> IS NOT ASSIGNED.
        add_missing( EXPORTING iv_tag = |table:{ lv_table }| CHANGING ct_missing = ct_missing ).
        CONTINUE.
      ENDIF.

      DATA(lv_out) = ``.
      LOOP AT <lt_lines> ASSIGNING FIELD-SYMBOL(<ls_line>).
        IF cl_abap_typedescr=>describe_by_data( <ls_line> )->kind <> cl_abap_typedescr=>kind_struct.
          add_missing( EXPORTING iv_tag = |table:{ lv_table } (dòng không phải cấu trúc)| CHANGING ct_missing = ct_missing ).
          EXIT.
        ENDIF.
        lv_out &&= fill_tags( EXPORTING iv_xml     = lv_row
                                        it_values  = to_values( <ls_line> )
                                        iv_prefix  = |table:{ lv_table }.|
                              CHANGING  ct_missing = ct_missing ).
      ENDLOOP.

      rv_xml = replace( val = rv_xml off = ls_row-offset len = ls_row-length with = lv_out ).
    ENDWHILE.
  ENDMETHOD.


  METHOD fill_tags.
    " iv_prefix rong: thay tag thuong; co prefix: chi thay tag bat dau bang prefix (dong bang)
    rv_xml = iv_xml.
    FIND ALL OCCURRENCES OF PCRE `\$\{([^}<]*)\}` IN iv_xml RESULTS DATA(lt_tag).

    DATA(lv_i) = lines( lt_tag ).
    WHILE lv_i > 0.
      DATA(ls_tag) = lt_tag[ lv_i ].
      lv_i -= 1.
      DATA(ls_sub) = ls_tag-submatches[ 1 ].
      DATA(lv_name) = condense( substring( val = iv_xml off = ls_sub-offset len = ls_sub-length ) ).

      IF iv_prefix IS NOT INITIAL.
        IF NOT lv_name CP |{ iv_prefix }*|.
          CONTINUE.
        ENDIF.
        lv_name = substring( val = lv_name off = strlen( iv_prefix ) ).
      ELSEIF lv_name CP 'table:*'.
        add_missing( EXPORTING iv_tag = |{ lv_name } (nằm ngoài dòng bảng)| CHANGING ct_missing = ct_missing ).
        CONTINUE.
      ENDIF.

      DATA(lv_value) = ``.
      READ TABLE it_values INTO DATA(ls_value) WITH TABLE KEY key = norm_key( lv_name ).
      IF sy-subrc = 0.
        lv_value = escape( val = ls_value-value format = cl_abap_format=>e_xml_text ).
      ELSE.
        add_missing( EXPORTING iv_tag = |{ iv_prefix }{ lv_name }| CHANGING ct_missing = ct_missing ).
      ENDIF.
      rv_xml = replace( val = rv_xml off = ls_tag-offset len = ls_tag-length with = lv_value ).
    ENDWHILE.
  ENDMETHOD.


  METHOD to_values.
    " Field thuong cua cau truc -> key chuan hoa (bo '_', viet hoa): doc_no khop ${docNo}
    DATA(lo_type) = cl_abap_typedescr=>describe_by_data( is_data ).
    IF lo_type->kind <> cl_abap_typedescr=>kind_struct.
      RETURN.
    ENDIF.

    DATA(lt_comp) = CAST cl_abap_structdescr( lo_type )->components.
    LOOP AT lt_comp INTO DATA(ls_comp).
      IF ls_comp-type_kind = cl_abap_typedescr=>typekind_table
         OR ls_comp-type_kind = cl_abap_typedescr=>typekind_struct1
         OR ls_comp-type_kind = cl_abap_typedescr=>typekind_struct2
         OR ls_comp-type_kind = cl_abap_typedescr=>typekind_oref
         OR ls_comp-type_kind = cl_abap_typedescr=>typekind_dref.
        CONTINUE.
      ENDIF.
      ASSIGN COMPONENT ls_comp-name OF STRUCTURE is_data TO FIELD-SYMBOL(<la_value>).
      INSERT VALUE #( key   = norm_key( CONV #( ls_comp-name ) )
                      value = to_text( <la_value> ) ) INTO TABLE rt_values.
    ENDLOOP.
  ENDMETHOD.


  METHOD to_text.
    CASE cl_abap_typedescr=>describe_by_data( ia_value )->type_kind.
      WHEN cl_abap_typedescr=>typekind_date.
        DATA(lv_date) = CONV d( ia_value ).
        IF lv_date IS NOT INITIAL.
          rv_text = |{ lv_date+6(2) }.{ lv_date+4(2) }.{ lv_date(4) }|.
        ENDIF.
      WHEN cl_abap_typedescr=>typekind_packed OR cl_abap_typedescr=>typekind_int
        OR cl_abap_typedescr=>typekind_int1 OR cl_abap_typedescr=>typekind_int2
        OR cl_abap_typedescr=>typekind_int8 OR cl_abap_typedescr=>typekind_float
        OR cl_abap_typedescr=>typekind_decfloat16 OR cl_abap_typedescr=>typekind_decfloat34.
        DATA(lv_number) = CONV decfloat34( ia_value ).
        rv_text = |{ lv_number NUMBER = USER }|.
      WHEN OTHERS.
        rv_text = |{ ia_value }|.
    ENDCASE.
  ENDMETHOD.


  METHOD norm_key.
    rv_key = to_upper( replace( val = iv_name sub = `_` with = `` occ = 0 ) ).
  ENDMETHOD.


  METHOD add_missing.
    IF NOT line_exists( ct_missing[ table_line = iv_tag ] ).
      APPEND iv_tag TO ct_missing.
    ENDIF.
  ENDMETHOD.

ENDCLASS.
