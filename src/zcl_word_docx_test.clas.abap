CLASS zcl_word_docx_test DEFINITION
  PUBLIC
  FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    CONSTANTS c_mime_docx TYPE string VALUE `application/vnd.openxmlformats-officedocument.wordprocessingml.document`.

    CONSTANTS: BEGIN OF c_msg_type,
                 error   TYPE c LENGTH 1 VALUE 'E',
                 warning TYPE c LENGTH 1 VALUE 'W',
                 success TYPE c LENGTH 1 VALUE 'S',
               END OF c_msg_type.

    TYPES: BEGIN OF ty_message,
             type TYPE c LENGTH 1,
             tag  TYPE string,
             text TYPE string,
           END OF ty_message,
           tt_message TYPE STANDARD TABLE OF ty_message WITH EMPTY KEY.

    CLASS-METHODS render
      IMPORTING iv_template TYPE xstring
                is_data     TYPE any
      EXPORTING ev_docx     TYPE xstring
                ev_error    TYPE string
                et_messages TYPE tt_message.

    CLASS-METHODS is_docx
      IMPORTING iv_content   TYPE xstring
      RETURNING VALUE(rv_ok) TYPE abap_bool.

    CLASS-METHODS number_to_words_vn
      IMPORTING iv_number      TYPE decfloat34
      RETURNING VALUE(rv_text) TYPE string.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_value,
             key   TYPE string,
             value TYPE string,
           END OF ty_value,
           tt_value TYPE HASHED TABLE OF ty_value WITH UNIQUE KEY key,
           tt_scope TYPE STANDARD TABLE OF REF TO data WITH EMPTY KEY,
           BEGIN OF ty_media,
             rid     TYPE string,
             target  TYPE string,
             content TYPE xstring,
           END OF ty_media,
           tt_media TYPE STANDARD TABLE OF ty_media WITH EMPTY KEY,
           BEGIN OF ty_row,
             offset TYPE i,
             length TYPE i,
             table  TYPE string,
           END OF ty_row,
           tt_row TYPE STANDARD TABLE OF ty_row WITH EMPTY KEY.

    CONSTANTS: c_emu_per_cm TYPE i VALUE 360000,
               c_png        TYPE x LENGTH 4 VALUE '89504E47',
               c_jpeg       TYPE x LENGTH 2 VALUE 'FFD8',
               c_break      TYPE string VALUE `</w:t><w:br/><w:t xml:space="preserve">`.

    DATA mo_zip     TYPE REF TO cl_abap_zip.
    DATA mt_missing  TYPE string_table.
    DATA mt_messages TYPE tt_message.
    DATA mt_media   TYPE tt_media.
    DATA mt_ext     TYPE string_table.
    DATA mv_image   TYPE i.

    METHODS run
      IMPORTING iv_template TYPE xstring
                is_data     TYPE any
      EXPORTING ev_docx     TYPE xstring
                ev_error    TYPE string
                et_messages TYPE tt_message.

    METHODS process_xml
      IMPORTING iv_xml        TYPE string
                it_scope      TYPE tt_scope
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS expand_blocks
      IMPORTING iv_xml        TYPE string
                it_scope      TYPE tt_scope
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS expand_rows
      IMPORTING iv_xml        TYPE string
                it_scope      TYPE tt_scope
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS put_images
      IMPORTING iv_xml        TYPE string
                it_scope      TYPE tt_scope
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS fill_tags
      IMPORTING iv_xml        TYPE string
                it_values     TYPE tt_value
                iv_prefix     TYPE string OPTIONAL
      RETURNING VALUE(rv_xml) TYPE string.

    METHODS flush_media
      IMPORTING iv_part TYPE string.

    METHODS add_content_types.

    METHODS add_missing
      IMPORTING iv_tag TYPE string.

    METHODS add_message
      IMPORTING iv_type TYPE c
                iv_tag  TYPE string
                iv_text TYPE string.

    METHODS get_text
      IMPORTING iv_name        TYPE string
      RETURNING VALUE(rv_text) TYPE string.

    METHODS put_text
      IMPORTING iv_name TYPE string
                iv_text TYPE string.

    CLASS-METHODS merge_split_tags
      IMPORTING iv_xml        TYPE string
      RETURNING VALUE(rv_xml) TYPE string.

    CLASS-METHODS merge_paragraph
      IMPORTING iv_par        TYPE string
      RETURNING VALUE(rv_par) TYPE string.

    CLASS-METHODS is_open
      IMPORTING iv_text        TYPE string
      RETURNING VALUE(rv_open) TYPE abap_bool.

    CLASS-METHODS paragraph_bounds
      IMPORTING iv_xml   TYPE string
                iv_pos   TYPE i
      EXPORTING ev_start TYPE i
                ev_end   TYPE i.

    CLASS-METHODS find_component
      IMPORTING it_scope      TYPE tt_scope
                iv_name       TYPE string
      RETURNING VALUE(rr_ref) TYPE REF TO data.

    CLASS-METHODS is_true
      IMPORTING it_scope       TYPE tt_scope
                iv_name        TYPE string
      RETURNING VALUE(rv_true) TYPE abap_bool.

    CLASS-METHODS scope_values
      IMPORTING it_scope         TYPE tt_scope
      RETURNING VALUE(rt_values) TYPE tt_value.

    CLASS-METHODS to_values
      IMPORTING is_data          TYPE any
      CHANGING  ct_values        TYPE tt_value.

    CLASS-METHODS to_text
      IMPORTING ia_value       TYPE any
      RETURNING VALUE(rv_text) TYPE string.

    CLASS-METHODS image_size
      IMPORTING iv_image  TYPE xstring
      EXPORTING ev_ext    TYPE string
                ev_width  TYPE i
                ev_height TYPE i.

    CLASS-METHODS read_triple
      IMPORTING iv_number      TYPE i
                iv_full        TYPE abap_bool
      RETURNING VALUE(rv_text) TYPE string.

    CLASS-METHODS norm_key
      IMPORTING iv_name       TYPE string
      RETURNING VALUE(rv_key) TYPE string.
ENDCLASS.



CLASS zcl_word_docx_test IMPLEMENTATION.

  METHOD render.
    " ${tag}, ${table:t.f} (dong bang, ${rowif:f}), ${#each:t}/${#if:f}/${#ifnot:f} (khoi doan van), ${image:f|cm}
    " et_messages: E = loi (khong sinh file), W = canh bao (van sinh), S = thanh cong
    CLEAR: ev_docx, ev_error, et_messages.
    IF is_docx( iv_template ) = abap_false.
      ev_error = `File template không phải .docx`.
      et_messages = VALUE #( ( type = c_msg_type-error text = ev_error ) ).
      RETURN.
    ENDIF.
    NEW zcl_word_docx_test( )->run( EXPORTING iv_template = iv_template
                                              is_data     = is_data
                                    IMPORTING ev_docx     = ev_docx
                                              ev_error    = ev_error
                                              et_messages = et_messages ).
  ENDMETHOD.


  METHOD is_docx.
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( EXPORTING  zip             = iv_content
                  EXCEPTIONS zip_parse_error = 1
                             OTHERS          = 2 ).
    rv_ok = xsdbool( sy-subrc = 0 AND line_exists( lo_zip->files[ name = `word/document.xml` ] ) ).
  ENDMETHOD.


  METHOD run.
    mo_zip = NEW cl_abap_zip( ).
    mo_zip->load( EXPORTING  zip             = iv_template
                  EXCEPTIONS zip_parse_error = 1
                             OTHERS          = 2 ).

    DATA lt_parts TYPE string_table.
    LOOP AT mo_zip->files INTO DATA(ls_file)
         WHERE name = `word/document.xml` OR name CP 'word/header*.xml' OR name CP 'word/footer*.xml'.
      APPEND ls_file-name TO lt_parts.
    ENDLOOP.

    DATA(lt_scope) = VALUE tt_scope( ( REF #( is_data ) ) ).

    TRY.
        LOOP AT lt_parts INTO DATA(lv_part).
          DATA(lv_xml) = merge_split_tags( get_text( lv_part ) ).
          put_text( iv_name = lv_part
                    iv_text = process_xml( iv_xml = lv_xml it_scope = lt_scope ) ).
          flush_media( lv_part ).
        ENDLOOP.
        add_content_types( ).
      CATCH cx_sy_conversion_codepage.
        ev_error = `Lỗi mã hóa ký tự trong template`.
        add_message( iv_type = c_msg_type-error iv_tag = `` iv_text = ev_error ).
        et_messages = mt_messages.
        RETURN.
    ENDTRY.

    IF mt_missing IS NOT INITIAL.
      ev_error = |Thiếu dữ liệu: { concat_lines_of( table = mt_missing sep = `, ` ) }|.
      et_messages = mt_messages.
      RETURN.
    ENDIF.
    ev_docx = mo_zip->save( ).
    add_message( iv_type = c_msg_type-success iv_tag = `` iv_text = |Đã tạo file ({ xstrlen( ev_docx ) } byte)| ).
    et_messages = mt_messages.
  ENDMETHOD.


  METHOD process_xml.
    rv_xml = expand_blocks( iv_xml = iv_xml it_scope = it_scope ).
    rv_xml = expand_rows( iv_xml = rv_xml it_scope = it_scope ).
    rv_xml = put_images( iv_xml = rv_xml it_scope = it_scope ).
    rv_xml = fill_tags( iv_xml = rv_xml it_values = scope_values( it_scope ) ).
  ENDMETHOD.


  METHOD expand_blocks.
    " Doan van chua ${#each:t} ... doan van chua ${/each:t}: lap khoi theo tung dong cua t (dong la pham vi du lieu moi)
    " ${#if:f} / ${#ifnot:f} ... ${/if:f}: giu hoac bo khoi; doan van chua the luon bi bo
    rv_xml = iv_xml.
    DO.
      FIND FIRST OCCURRENCE OF PCRE `\$\{\#(each|if|ifnot):(\w+)\}` IN rv_xml
           MATCH OFFSET DATA(lv_pos) SUBMATCHES DATA(lv_kind) DATA(lv_name).
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.

      paragraph_bounds( EXPORTING iv_xml = rv_xml iv_pos = lv_pos
                        IMPORTING ev_start = DATA(lv_start) ev_end = DATA(lv_inner_start) ).

      DATA(lv_close) = |$\{/{ COND string( WHEN lv_kind = `ifnot` THEN `if` ELSE lv_kind ) }:{ lv_name }\}|.
      DATA(lv_close_pos) = find( val = rv_xml sub = lv_close off = lv_inner_start ).
      IF lv_close_pos < 0.
        add_missing( |#{ lv_kind }:{ lv_name } (thiếu { lv_close })| ).
        rv_xml = replace( val = rv_xml off = lv_start len = lv_inner_start - lv_start with = `` ).
        CONTINUE.
      ENDIF.
      paragraph_bounds( EXPORTING iv_xml = rv_xml iv_pos = lv_close_pos
                        IMPORTING ev_start = DATA(lv_inner_end) ev_end = DATA(lv_end) ).

      DATA(lv_inner) = substring( val = rv_xml off = lv_inner_start len = lv_inner_end - lv_inner_start ).
      DATA(lv_out) = ``.
      CASE lv_kind.
        WHEN `each`.
          DATA(lr_table) = find_component( it_scope = it_scope iv_name = lv_name ).
          FIELD-SYMBOLS <lt_lines> TYPE ANY TABLE.
          UNASSIGN <lt_lines>.
          IF lr_table IS BOUND.
            ASSIGN lr_table->* TO FIELD-SYMBOL(<la_table>).
            IF cl_abap_typedescr=>describe_by_data( <la_table> )->kind = cl_abap_typedescr=>kind_table.
              ASSIGN lr_table->* TO <lt_lines>.
            ENDIF.
          ENDIF.
          IF <lt_lines> IS NOT ASSIGNED.
            add_missing( |#each:{ lv_name }| ).
          ELSE.
            IF lines( <lt_lines> ) = 0.
              add_message( iv_type = c_msg_type-warning iv_tag = |#each:{ lv_name }| iv_text = |Bảng { lv_name } rỗng, bỏ khối lặp| ).
            ENDIF.
            LOOP AT <lt_lines> ASSIGNING FIELD-SYMBOL(<ls_line>).
              lv_out &&= process_xml( iv_xml   = lv_inner
                                      it_scope = VALUE #( ( REF #( <ls_line> ) ) ( LINES OF it_scope ) ) ).
            ENDLOOP.
          ENDIF.
        WHEN `if` OR `ifnot`.
          IF find_component( it_scope = it_scope iv_name = lv_name ) IS NOT BOUND.
            add_missing( |#{ lv_kind }:{ lv_name }| ).
          ENDIF.
          IF lv_kind = `if`.
            lv_out = COND #( WHEN is_true( it_scope = it_scope iv_name = lv_name ) = abap_true THEN lv_inner ).
          ELSE.
            lv_out = COND #( WHEN is_true( it_scope = it_scope iv_name = lv_name ) = abap_false THEN lv_inner ).
          ENDIF.
      ENDCASE.
      rv_xml = replace( val = rv_xml off = lv_start len = lv_end - lv_start with = lv_out ).
    ENDDO.
  ENDMETHOD.


  METHOD expand_rows.
    " Cac dong bang lien nhau cung ${table:t.} = 1 khoi, lap theo tung dong cua t; ${rowif:f}/${rowifnot:f} chon dong mau
    rv_xml = iv_xml.
    FIND ALL OCCURRENCES OF PCRE `(?s)<w:tr(?:\s[^>]*)?>.*?</w:tr>` IN iv_xml RESULTS DATA(lt_match).

    DATA lt_rows TYPE tt_row.
    LOOP AT lt_match INTO DATA(ls_match).
      DATA(lv_row) = substring( val = iv_xml off = ls_match-offset len = ls_match-length ).
      FIND FIRST OCCURRENCE OF PCRE `\$\{table:(\w+)\.` IN lv_row SUBMATCHES DATA(lv_table).
      IF sy-subrc = 0.
        APPEND VALUE #( offset = ls_match-offset length = ls_match-length table = lv_table ) TO lt_rows.
      ENDIF.
    ENDLOOP.

    DATA(lv_i) = lines( lt_rows ).
    WHILE lv_i > 0.
      DATA(lv_last) = lv_i.
      DATA(lv_first) = lv_i.
      WHILE lv_first > 1
        AND lt_rows[ lv_first - 1 ]-table = lt_rows[ lv_first ]-table
        AND lt_rows[ lv_first - 1 ]-offset + lt_rows[ lv_first - 1 ]-length = lt_rows[ lv_first ]-offset.
        lv_first -= 1.
      ENDWHILE.
      lv_i = lv_first - 1.

      DATA(ls_first) = lt_rows[ lv_first ].
      DATA(ls_last)  = lt_rows[ lv_last ].
      DATA(lr_table) = find_component( it_scope = it_scope iv_name = ls_first-table ).
      FIELD-SYMBOLS <lt_lines> TYPE ANY TABLE.
      UNASSIGN <lt_lines>.
      IF lr_table IS BOUND.
        ASSIGN lr_table->* TO FIELD-SYMBOL(<la_table>).
        IF cl_abap_typedescr=>describe_by_data( <la_table> )->kind = cl_abap_typedescr=>kind_table.
          ASSIGN lr_table->* TO <lt_lines>.
        ENDIF.
      ENDIF.
      IF <lt_lines> IS NOT ASSIGNED.
        add_missing( |table:{ ls_first-table }| ).
        CONTINUE.
      ENDIF.

      DATA(lv_out) = ``.
      IF lines( <lt_lines> ) = 0.
        add_message( iv_type = c_msg_type-warning iv_tag = |table:{ ls_first-table }| iv_text = |Bảng { ls_first-table } rỗng, bỏ dòng bảng| ).
      ENDIF.
      LOOP AT <lt_lines> ASSIGNING FIELD-SYMBOL(<ls_line>).
        DATA(lt_line_scope) = VALUE tt_scope( ( REF #( <ls_line> ) ) ).
        DATA(lt_values) = scope_values( lt_line_scope ).
        LOOP AT lt_rows INTO DATA(ls_row) FROM lv_first TO lv_last.
          DATA(lv_tpl) = substring( val = iv_xml off = ls_row-offset len = ls_row-length ).
          FIND FIRST OCCURRENCE OF PCRE `\$\{rowif(not)?:(\w+)\}` IN lv_tpl SUBMATCHES DATA(lv_not) DATA(lv_field).
          IF sy-subrc = 0.
            IF find_component( it_scope = lt_line_scope iv_name = lv_field ) IS NOT BOUND.
              add_missing( |rowif:{ lv_field }| ).
            ENDIF.
            DATA(lv_true) = is_true( it_scope = lt_line_scope iv_name = lv_field ).
            IF ( lv_not IS INITIAL AND lv_true = abap_false ) OR ( lv_not IS NOT INITIAL AND lv_true = abap_true ).
              CONTINUE.
            ENDIF.
            lv_tpl = replace( val = lv_tpl pcre = `\$\{rowif(not)?:\w+\}` with = `` occ = 0 ).
          ENDIF.
          lv_out &&= fill_tags( iv_xml = lv_tpl it_values = lt_values iv_prefix = |table:{ ls_first-table }.| ).
        ENDLOOP.
      ENDLOOP.

      rv_xml = replace( val  = rv_xml
                        off  = ls_first-offset
                        len  = ls_last-offset + ls_last-length - ls_first-offset
                        with = lv_out ).
    ENDWHILE.
  ENDMETHOD.


  METHOD put_images.
    " ${image:f} hoac ${image:f|4} (rong 4 cm, cao theo ty le anh); f la xstring PNG/JPEG, rong -> bo tag
    rv_xml = iv_xml.
    FIND ALL OCCURRENCES OF PCRE `\$\{image:(\w+)(?:\|([\d.]+))?\}` IN iv_xml RESULTS DATA(lt_tag).

    DATA(lv_i) = lines( lt_tag ).
    WHILE lv_i > 0.
      DATA(ls_tag) = lt_tag[ lv_i ].
      lv_i -= 1.
      DATA(ls_name) = ls_tag-submatches[ 1 ].
      DATA(lv_name) = substring( val = iv_xml off = ls_name-offset len = ls_name-length ).
      DATA(lv_width_cm) = CONV decfloat34( 4 ).
      DATA(ls_width) = ls_tag-submatches[ 2 ].
      IF ls_width-offset >= 0 AND ls_width-length > 0.
        lv_width_cm = substring( val = iv_xml off = ls_width-offset len = ls_width-length ).
      ENDIF.

      DATA(lv_run) = ``.
      DATA(lr_image) = find_component( it_scope = it_scope iv_name = lv_name ).
      IF lr_image IS NOT BOUND.
        add_missing( |image:{ lv_name }| ).
      ELSE.
        ASSIGN lr_image->* TO FIELD-SYMBOL(<lv_image>).
        DATA(lv_content) = CONV xstring( <lv_image> ).
        IF lv_content IS INITIAL.
          add_message( iv_type = c_msg_type-warning iv_tag = |image:{ lv_name }| iv_text = |Ảnh { lv_name } rỗng, bỏ ảnh| ).
        ELSE.
          image_size( EXPORTING iv_image  = lv_content
                      IMPORTING ev_ext    = DATA(lv_ext)
                                ev_width  = DATA(lv_px_w)
                                ev_height = DATA(lv_px_h) ).
          IF lv_ext IS INITIAL.
            add_missing( |image:{ lv_name } (không phải PNG/JPEG)| ).
          ELSE.
            mv_image += 1.
            DATA(lv_id)  = 9000 + mv_image.
            DATA(lv_rid) = |rIdZimg{ mv_image }|.
            DATA(lv_cx)  = CONV i( lv_width_cm * c_emu_per_cm ).
            DATA(lv_cy)  = COND i( WHEN lv_px_w > 0 AND lv_px_h > 0 THEN CONV decfloat34( lv_cx ) * lv_px_h / lv_px_w ELSE lv_cx ).
            APPEND VALUE #( rid = lv_rid target = |media/zimg{ mv_image }.{ lv_ext }| content = lv_content ) TO mt_media.
            IF NOT line_exists( mt_ext[ table_line = lv_ext ] ).
              APPEND lv_ext TO mt_ext.
            ENDIF.
            lv_run = |</w:t></w:r><w:r><w:drawing>|
                  && |<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">|
                  && |<wp:extent cx="{ lv_cx }" cy="{ lv_cy }"/><wp:docPr id="{ lv_id }" name="Picture { lv_id }"/>|
                  && |<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">|
                  && |<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">|
                  && |<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">|
                  && |<pic:nvPicPr><pic:cNvPr id="{ lv_id }" name="zimg{ mv_image }.{ lv_ext }"/><pic:cNvPicPr/></pic:nvPicPr>|
                  && |<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{ lv_rid }"/>|
                  && |<a:stretch><a:fillRect/></a:stretch></pic:blipFill>|
                  && |<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{ lv_cx }" cy="{ lv_cy }"/></a:xfrm>|
                  && |<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>|
                  && |</a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t xml:space="preserve">|.
          ENDIF.
        ENDIF.
      ENDIF.
      rv_xml = replace( val = rv_xml off = ls_tag-offset len = ls_tag-length with = lv_run ).
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
        add_missing( |{ lv_name } (nằm ngoài dòng bảng)| ).
        CONTINUE.
      ENDIF.

      DATA(lv_value) = ``.
      READ TABLE it_values INTO DATA(ls_value) WITH TABLE KEY key = norm_key( lv_name ).
      IF sy-subrc = 0.
        lv_value = escape( val = ls_value-value format = cl_abap_format=>e_xml_text ).
        lv_value = replace( val = lv_value sub = cl_abap_char_utilities=>cr_lf with = cl_abap_char_utilities=>newline occ = 0 ).
        lv_value = replace( val = lv_value sub = cl_abap_char_utilities=>newline with = c_break occ = 0 ).
      ELSE.
        add_missing( |{ iv_prefix }{ lv_name }| ).
      ENDIF.
      rv_xml = replace( val = rv_xml off = ls_tag-offset len = ls_tag-length with = lv_value ).
    ENDWHILE.
  ENDMETHOD.


  METHOD flush_media.
    " Anh vua chen vao part -> word/media/* + quan he trong word/_rels/<part>.rels
    IF mt_media IS INITIAL.
      RETURN.
    ENDIF.

    DATA(lv_rels_name) = |word/_rels/{ substring_after( val = iv_part sub = `word/` ) }.rels|.
    DATA(lv_rels) = COND string( WHEN line_exists( mo_zip->files[ name = lv_rels_name ] ) THEN get_text( lv_rels_name ) ).
    IF lv_rels IS INITIAL.
      lv_rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
             && `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`.
    ENDIF.

    DATA(lv_new) = ``.
    LOOP AT mt_media INTO DATA(ls_media).
      lv_new &&= |<Relationship Id="{ ls_media-rid }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{ ls_media-target }"/>|.
      mo_zip->add( name = |word/{ ls_media-target }| content = ls_media-content ).
    ENDLOOP.
    put_text( iv_name = lv_rels_name
              iv_text = replace( val = lv_rels sub = `</Relationships>` with = lv_new && `</Relationships>` ) ).
    CLEAR mt_media.
  ENDMETHOD.


  METHOD add_content_types.
    IF mt_ext IS INITIAL.
      RETURN.
    ENDIF.
    DATA(lv_types) = get_text( `[Content_Types].xml` ).
    LOOP AT mt_ext INTO DATA(lv_ext).
      FIND |Extension="{ lv_ext }"| IN lv_types IGNORING CASE.
      IF sy-subrc <> 0.
        lv_types = replace( val  = lv_types
                            sub  = `</Types>`
                            with = |<Default Extension="{ lv_ext }" ContentType="image/{ lv_ext }"/></Types>| ).
      ENDIF.
    ENDLOOP.
    put_text( iv_name = `[Content_Types].xml` iv_text = lv_types ).
  ENDMETHOD.


  METHOD add_missing.
    IF NOT line_exists( mt_missing[ table_line = iv_tag ] ).
      APPEND iv_tag TO mt_missing.
      add_message( iv_type = c_msg_type-error iv_tag = iv_tag iv_text = |Thiếu dữ liệu: { iv_tag }| ).
    ENDIF.
  ENDMETHOD.


  METHOD add_message.
    IF NOT line_exists( mt_messages[ type = iv_type tag = iv_tag text = iv_text ] ).
      APPEND VALUE #( type = iv_type tag = iv_tag text = iv_text ) TO mt_messages.
    ENDIF.
  ENDMETHOD.


  METHOD get_text.
    DATA lv_raw TYPE xstring.
    mo_zip->get( EXPORTING  name                    = iv_name
                 IMPORTING  content                 = lv_raw
                 EXCEPTIONS zip_index_error         = 1
                            zip_decompression_error = 2
                            OTHERS                  = 3 ).
    IF sy-subrc = 0.
      rv_text = cl_abap_conv_codepage=>create_in( )->convert( lv_raw ).
    ENDIF.
  ENDMETHOD.


  METHOD put_text.
    mo_zip->delete( EXPORTING  name            = iv_name
                    EXCEPTIONS zip_index_error = 1
                               OTHERS          = 2 ).
    mo_zip->add( name    = iv_name
                 content = cl_abap_conv_codepage=>create_out( )->convert( iv_text ) ).
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


  METHOD paragraph_bounds.
    " Doan van <w:p ...>...</w:p> chua vi tri iv_pos
    FIND ALL OCCURRENCES OF PCRE `<w:p(?:\s[^>]*)?(?<!/)>` IN SECTION LENGTH iv_pos OF iv_xml RESULTS DATA(lt_open).
    ev_start = COND #( WHEN lt_open IS NOT INITIAL THEN lt_open[ lines( lt_open ) ]-offset ELSE iv_pos ).
    DATA(lv_close) = find( val = iv_xml sub = `</w:p>` off = iv_pos ).
    ev_end = COND #( WHEN lv_close >= 0 THEN lv_close + 6 ELSE strlen( iv_xml ) ).
  ENDMETHOD.


  METHOD find_component.
    " Tim component theo ten chuan hoa, tu pham vi trong (dong dang lap) ra ngoai (du lieu goc)
    LOOP AT it_scope INTO DATA(lr_scope).
      ASSIGN lr_scope->* TO FIELD-SYMBOL(<ls_scope>).
      DATA(lo_type) = cl_abap_typedescr=>describe_by_data( <ls_scope> ).
      IF lo_type->kind <> cl_abap_typedescr=>kind_struct.
        CONTINUE.
      ENDIF.
      DATA(lt_comp) = CAST cl_abap_structdescr( lo_type )->components.
      LOOP AT lt_comp INTO DATA(ls_comp).
        IF norm_key( CONV #( ls_comp-name ) ) = norm_key( iv_name ).
          ASSIGN COMPONENT ls_comp-name OF STRUCTURE <ls_scope> TO FIELD-SYMBOL(<la_value>).
          rr_ref = REF #( <la_value> ).
          RETURN.
        ENDIF.
      ENDLOOP.
    ENDLOOP.
  ENDMETHOD.


  METHOD is_true.
    " Bang: co dong; field: khac rong / khac 0
    DATA(lr_ref) = find_component( it_scope = it_scope iv_name = iv_name ).
    IF lr_ref IS NOT BOUND.
      RETURN.
    ENDIF.
    ASSIGN lr_ref->* TO FIELD-SYMBOL(<la_value>).
    IF cl_abap_typedescr=>describe_by_data( <la_value> )->kind = cl_abap_typedescr=>kind_table.
      FIELD-SYMBOLS <lt_any> TYPE ANY TABLE.
      ASSIGN lr_ref->* TO <lt_any>.
      rv_true = xsdbool( lines( <lt_any> ) > 0 ).
    ELSE.
      rv_true = xsdbool( <la_value> IS NOT INITIAL ).
    ENDIF.
  ENDMETHOD.


  METHOD scope_values.
    LOOP AT it_scope INTO DATA(lr_scope).
      ASSIGN lr_scope->* TO FIELD-SYMBOL(<ls_scope>).
      to_values( EXPORTING is_data = <ls_scope> CHANGING ct_values = rt_values ).
    ENDLOOP.
  ENDMETHOD.


  METHOD to_values.
    " Field thuong cua cau truc -> key chuan hoa (bo '_', viet hoa): doc_no khop ${docNo}; key da co (pham vi trong) giu nguyen
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
         OR ls_comp-type_kind = cl_abap_typedescr=>typekind_dref
         OR ls_comp-type_kind = cl_abap_typedescr=>typekind_xstring.
        CONTINUE.
      ENDIF.
      ASSIGN COMPONENT ls_comp-name OF STRUCTURE is_data TO FIELD-SYMBOL(<la_value>).
      INSERT VALUE #( key   = norm_key( CONV #( ls_comp-name ) )
                      value = to_text( <la_value> ) ) INTO TABLE ct_values.
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


  METHOD image_size.
    " Kich thuoc pixel: PNG doc IHDR, JPEG tim marker SOF0/1/2
    CLEAR: ev_ext, ev_width, ev_height.
    DATA lv_x4 TYPE x LENGTH 4.
    DATA lv_x2 TYPE x LENGTH 2.
    DATA lv_x1 TYPE x LENGTH 1.
    DATA(lv_len) = xstrlen( iv_image ).

    IF lv_len >= 24 AND iv_image(4) = c_png.
      ev_ext = `png`.
      lv_x4 = iv_image+16(4).
      ev_width = lv_x4.
      lv_x4 = iv_image+20(4).
      ev_height = lv_x4.
      RETURN.
    ENDIF.

    IF lv_len < 4 OR iv_image(2) <> c_jpeg.
      RETURN.
    ENDIF.
    ev_ext = `jpeg`.
    DATA(lv_off) = 2.
    WHILE lv_off + 9 < lv_len.
      lv_x1 = iv_image+lv_off(1).
      IF lv_x1 <> 'FF'.
        RETURN.
      ENDIF.
      DATA(lv_marker_off) = lv_off + 1.
      lv_x1 = iv_image+lv_marker_off(1).
      IF lv_x1 = 'C0' OR lv_x1 = 'C1' OR lv_x1 = 'C2'.
        DATA(lv_h_off) = lv_off + 5.
        DATA(lv_w_off) = lv_off + 7.
        lv_x2 = iv_image+lv_h_off(2).
        ev_height = lv_x2.
        lv_x2 = iv_image+lv_w_off(2).
        ev_width = lv_x2.
        RETURN.
      ENDIF.
      DATA(lv_seg_off) = lv_off + 2.
      lv_x2 = iv_image+lv_seg_off(2).
      DATA(lv_seg_len) = CONV i( lv_x2 ).
      lv_off = lv_off + 2 + lv_seg_len.
    ENDWHILE.
  ENDMETHOD.


  METHOD number_to_words_vn.
    " Phan nguyen, doc theo nhom 3 chu so: 2.640.000 -> Hai triệu sáu trăm bốn mươi nghìn
    DATA(lv_value) = CONV int8( trunc( abs( iv_number ) ) ).
    IF lv_value = 0.
      rv_text = `Không`.
      RETURN.
    ENDIF.

    DATA lt_group TYPE STANDARD TABLE OF i WITH EMPTY KEY.
    WHILE lv_value > 0.
      APPEND CONV i( lv_value MOD 1000 ) TO lt_group.
      lv_value = lv_value DIV 1000.
    ENDWHILE.

    DATA(lt_unit) = VALUE string_table( ( `` ) ( `nghìn` ) ( `triệu` ) ( `tỷ` ) ( `nghìn tỷ` ) ( `triệu tỷ` ) ( `tỷ tỷ` ) ).
    DATA(lv_top) = lines( lt_group ).
    DATA(lv_i) = lv_top.
    WHILE lv_i > 0.
      DATA(lv_group) = lt_group[ lv_i ].
      IF lv_group > 0.
        DATA(lv_part) = read_triple( iv_number = lv_group iv_full = xsdbool( lv_i < lv_top ) ).
        DATA(lv_unit) = lt_unit[ lv_i ].
        rv_text = condense( |{ rv_text } { lv_part } { lv_unit }| ).
      ENDIF.
      lv_i -= 1.
    ENDWHILE.

    IF iv_number < 0.
      rv_text = |âm { rv_text }|.
    ENDIF.
    rv_text = to_upper( substring( val = rv_text len = 1 ) ) && substring( val = rv_text off = 1 ).
  ENDMETHOD.


  METHOD read_triple.
    DATA(lt_digit) = VALUE string_table( ( `không` ) ( `một` ) ( `hai` ) ( `ba` ) ( `bốn` )
                                         ( `năm` ) ( `sáu` ) ( `bảy` ) ( `tám` ) ( `chín` ) ).
    DATA(lv_hundred) = iv_number DIV 100.
    DATA(lv_ten)     = ( iv_number MOD 100 ) DIV 10.
    DATA(lv_one)     = iv_number MOD 10.

    DATA lt_word TYPE string_table.
    IF lv_hundred > 0 OR iv_full = abap_true.
      APPEND |{ lt_digit[ lv_hundred + 1 ] } trăm| TO lt_word.
    ENDIF.

    CASE lv_ten.
      WHEN 0.
        IF lv_one > 0 AND ( lv_hundred > 0 OR iv_full = abap_true ).
          APPEND `lẻ` TO lt_word.
        ENDIF.
      WHEN 1.
        APPEND `mười` TO lt_word.
      WHEN OTHERS.
        APPEND |{ lt_digit[ lv_ten + 1 ] } mươi| TO lt_word.
    ENDCASE.

    IF lv_one > 0.
      APPEND COND string( WHEN lv_one = 1 AND lv_ten >= 2 THEN `mốt`
                          WHEN lv_one = 5 AND lv_ten >= 1 THEN `lăm`
                          WHEN lv_one = 4 AND lv_ten >= 2 THEN `tư`
                          ELSE lt_digit[ lv_one + 1 ] ) TO lt_word.
    ENDIF.
    rv_text = concat_lines_of( table = lt_word sep = ` ` ).
  ENDMETHOD.


  METHOD norm_key.
    rv_key = to_upper( replace( val = iv_name sub = `_` with = `` occ = 0 ) ).
  ENDMETHOD.

ENDCLASS.
