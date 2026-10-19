CLASS lhc_file DEFINITION INHERITING FROM cl_abap_behavior_handler.
  PRIVATE SECTION.
    METHODS get_global_authorizations FOR GLOBAL AUTHORIZATION
      IMPORTING REQUEST requested_authorizations FOR File RESULT result.

    CONSTANTS c_mime_xlsx TYPE string VALUE `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.

    METHODS setVersion FOR DETERMINE ON SAVE
      IMPORTING keys FOR File~setVersion.

    METHODS checkTemplate FOR VALIDATE ON SAVE
      IMPORTING keys FOR File~checkTemplate.

    METHODS is_bad_template
      IMPORTING iv_content    TYPE xstring
      RETURNING VALUE(rv_bad) TYPE abap_bool.
ENDCLASS.



CLASS lhc_file IMPLEMENTATION.

  METHOD get_global_authorizations.
    result-%create = if_abap_behv=>auth-allowed.
    result-%update = if_abap_behv=>auth-allowed.
    result-%delete = if_abap_behv=>auth-allowed.
  ENDMETHOD.


  METHOD setVersion.
    " Moi lan upload cung FileId = 1 version moi (max + 1)
    READ ENTITIES OF zr_xl_file IN LOCAL MODE
      ENTITY File
        FIELDS ( FileId Version )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_file).

    DELETE lt_file WHERE Version IS NOT INITIAL.
    IF lt_file IS INITIAL.
      RETURN.
    ENDIF.

    LOOP AT lt_file ASSIGNING FIELD-SYMBOL(<ls_file>).
      SELECT MAX( version ) FROM zxl_file
        WHERE file_id = @<ls_file>-FileId
        INTO @DATA(lv_max).
      <ls_file>-Version = lv_max + 1.
    ENDLOOP.

    MODIFY ENTITIES OF zr_xl_file IN LOCAL MODE
      ENTITY File
        UPDATE FIELDS ( Version )
        WITH VALUE #( FOR ls IN lt_file ( %tky = ls-%tky Version = ls-Version ) ).
  ENDMETHOD.


  METHOD checkTemplate.
    " Chan template co phan ExcelJS lam mat hoac doc loi
    READ ENTITIES OF zr_xl_file IN LOCAL MODE
      ENTITY File
        FIELDS ( FileContent MimeType )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_file).

    LOOP AT lt_file INTO DATA(ls_file).
      APPEND VALUE #( %tky = ls_file-%tky %state_area = 'CHECK_TEMPLATE' ) TO reported-file.
      IF ls_file-MimeType <> c_mime_xlsx.
        CONTINUE.
      ENDIF.

      IF is_bad_template( ls_file-FileContent ) = abap_false.
        CONTINUE.
      ENDIF.
      APPEND VALUE #( %tky = ls_file-%tky ) TO failed-file.
      APPEND VALUE #( %tky                 = ls_file-%tky
                      %state_area          = 'CHECK_TEMPLATE'
                      %element-FileContent = if_abap_behv=>mk-on
                      %msg                 = new_message_with_text( severity = if_abap_behv_message=>severity-error
                                                                    text     = `Lỗi template` ) ) TO reported-file.
    ENDLOOP.
  ENDMETHOD.


  METHOD is_bad_template.
    " Khong phai xlsx, hoac co chart / pivot / slicer / timeline / anh chen san
    DATA(lo_zip) = NEW cl_abap_zip( ).
    lo_zip->load( EXPORTING  zip             = iv_content
                  EXCEPTIONS zip_parse_error = 1
                             OTHERS          = 2 ).
    IF sy-subrc <> 0 OR NOT line_exists( lo_zip->files[ name = `xl/workbook.xml` ] ).
      rv_bad = abap_true.
      RETURN.
    ENDIF.

    LOOP AT lo_zip->files TRANSPORTING NO FIELDS
         WHERE name CP 'xl/charts/*' OR name CP 'xl/pivotTables/*' OR name CP 'xl/pivotCache/*'
            OR name CP 'xl/slicer*' OR name CP 'xl/timeline*' OR name CP 'xl/media/*'.
      rv_bad = abap_true.
      RETURN.
    ENDLOOP.
  ENDMETHOD.

ENDCLASS.
