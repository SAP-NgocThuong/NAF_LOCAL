CLASS lhc_template DEFINITION INHERITING FROM cl_abap_behavior_handler.
  PRIVATE SECTION.
    METHODS get_global_authorizations FOR GLOBAL AUTHORIZATION
      IMPORTING REQUEST requested_authorizations FOR Template RESULT result.

    METHODS checkTemplate FOR VALIDATE ON SAVE
      IMPORTING keys FOR Template~checkTemplate.
ENDCLASS.



CLASS lhc_template IMPLEMENTATION.

  METHOD get_global_authorizations.
    result-%create = if_abap_behv=>auth-allowed.
    result-%update = if_abap_behv=>auth-allowed.
    result-%delete = if_abap_behv=>auth-allowed.
  ENDMETHOD.


  METHOD checkTemplate.
    READ ENTITIES OF zr_word_tmpl_test IN LOCAL MODE
      ENTITY Template
        FIELDS ( TmplContent TmplMime )
        WITH CORRESPONDING #( keys )
      RESULT DATA(lt_tmpl).

    LOOP AT lt_tmpl INTO DATA(ls_tmpl).
      APPEND VALUE #( %tky = ls_tmpl-%tky %state_area = 'CHECK_TEMPLATE' ) TO reported-template.
      IF ls_tmpl-TmplContent IS INITIAL OR ls_tmpl-TmplMime <> zcl_word_docx_test=>c_mime_docx
         OR zcl_word_docx_test=>is_docx( ls_tmpl-TmplContent ) = abap_true.
        CONTINUE.
      ENDIF.
      APPEND VALUE #( %tky = ls_tmpl-%tky ) TO failed-template.
      APPEND VALUE #( %tky                 = ls_tmpl-%tky
                      %state_area          = 'CHECK_TEMPLATE'
                      %element-TmplContent = if_abap_behv=>mk-on
                      %msg                 = new_message_with_text( severity = if_abap_behv_message=>severity-error
                                                                    text     = `File không phải .docx` ) ) TO reported-template.
    ENDLOOP.
  ENDMETHOD.

ENDCLASS.
