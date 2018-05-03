CLASS zclgl_report_docx DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF gtys_imgs,
        id          TYPE zifgl_reports_img=>gtyv_id,
        height      TYPE tdhghtpix,
        width       TYPE tdwidthpix,
        percent     TYPE numc3,
        size_in_doc TYPE abap_bool,
        bmp         TYPE xstring,
        id_in_doc   TYPE string,
*        i_img       TYPE REF TO zifgl_reports_img,
        o_imagepart TYPE REF TO cl_oxml_imagepart,
      END OF gtys_imgs .
    TYPES:
      BEGIN OF gtys_tables,
        tpath TYPE  char255,
        lines TYPE sytabix,
      END OF gtys_tables .
    TYPES:
      BEGIN OF gtys_values,
        fpath TYPE char255,
        value TYPE string,
      END OF gtys_values .
    TYPES:
      gtyth_imgs   TYPE HASHED TABLE OF gtys_imgs WITH UNIQUE KEY id .
    TYPES:
      gtyth_tables TYPE HASHED TABLE OF gtys_tables WITH UNIQUE KEY tpath .
    TYPES:
      gtyth_values TYPE HASHED TABLE OF gtys_values WITH UNIQUE KEY fpath .
    DATA gv_dummy TYPE text256.
    CONSTANTS gc_hyperlink_tab TYPE char1 VALUE cl_abap_char_utilities=>vertical_tab ##NO_TEXT.

    CLASS-METHODS create_hyperlink
      IMPORTING
        !iv_text       TYPE clike
        !iv_url        TYPE clike
      RETURNING
        VALUE(rv_href) TYPE string .
    CLASS-METHODS get_hyperlink
      IMPORTING
        !iv_href TYPE clike
      EXPORTING
        !ev_text TYPE clike
        !ev_url  TYPE clike .
    METHODS run
      IMPORTING
        VALUE(iv_template)   TYPE zegl_report_template
        VALUE(iv_protect)    TYPE abap_bool DEFAULT abap_false
        !iv_filename         TYPE clike OPTIONAL
        VALUE(iv_debug_mode) TYPE abap_bool DEFAULT abap_false
        !iv_open             TYPE abap_bool DEFAULT abap_true
        !ith_imgs            TYPE gtyth_imgs OPTIONAL
        !ith_tables          TYPE gtyth_tables OPTIONAL
        !ith_values          TYPE gtyth_values OPTIONAL
      EXPORTING
        !ev_content          TYPE xstring
      RAISING
        zcx_gl_report_docx .
protected section.

  methods DEBUG .
  methods DOWNLOAD
    importing
      !IV_FILENAME type CLIKE
      !IV_OPEN type ABAP_BOOL default ABAP_TRUE
    raising
      ZCX_GL_REPORT_DOCX .
  methods FINAL_CHECK
    raising
      ZCX_GL_REPORT_DOCX .
  methods OPEN_TEMPLATE
    importing
      !IV_TEMPLATE type CLIKE
    raising
      ZCX_GL_REPORT_DOCX .
  PRIVATE SECTION.

    CONSTANTS gc_block_name TYPE string VALUE 'block_name' ##NO_TEXT.
    DATA gi_snode TYPE REF TO if_sxml_node .
    DATA gi_sreader TYPE REF TO if_sxml_reader .
    DATA gi_svalue_node TYPE REF TO if_sxml_value_node .
    DATA gi_swriter TYPE REF TO if_sxml_writer .
    DATA go_blok TYPE REF TO zclgl_report_docx_block_xml .
    DATA go_doc TYPE REF TO cl_docx_document .
    DATA go_documentpart TYPE REF TO cl_docx_maindocumentpart .
    DATA go_file TYPE REF TO zclgl_report_docx_file_xml .
    DATA gth_imgs TYPE gtyth_imgs .
    DATA gth_tables TYPE gtyth_tables .
    DATA gth_values TYPE gtyth_values .
    DATA gv_debug_mode TYPE abap_bool .
    DATA gv_document TYPE xstring .
    DATA gv_final_doc TYPE xstring .

    METHODS _clean .
    METHODS _find_variable
      IMPORTING
        !iv_tab_pref TYPE char255
        !ii_sreader  LIKE gi_sreader
      RAISING
        zcx_gl_report_docx .
    METHODS _find_variable_hyperlink
      CHANGING
        !cv_value TYPE clike .
    METHODS _find_variable_img
      IMPORTING
        !is_img     TYPE gtys_imgs
        !ii_sreader LIKE gi_sreader .
    METHODS _find_variable_prep_val
      CHANGING
        !cv_str TYPE string
      RAISING
        zcx_gl_report_docx .
    METHODS _finish_seq_access .
    METHODS _from_blocks_to_doc
      IMPORTING
        !io_blok           LIKE go_blok
        VALUE(iv_tab_pref) TYPE char255 OPTIONAL
      RAISING
        zcx_gl_report_docx .
    METHODS _get_nsuri_by_prefix
      IMPORTING
        !iv_prefix     TYPE string
      RETURNING
        VALUE(rv_nuri) TYPE string .
    METHODS _insert_img
      RAISING
        zcx_gl_report_docx .
    METHODS _prep_seq_access
      IMPORTING
        !iv_part TYPE xstring .
    METHODS _refresh_global_data
      RAISING
        zcx_gl_report_docx .
    METHODS _set_attribute_for_first_eleme
      IMPORTING
        !ii_element TYPE REF TO if_sxml_open_element .
    METHODS _set_blok
      IMPORTING
        !io_blok LIKE go_blok
        !iv_name TYPE clike .
    METHODS _set_value
      IMPORTING
        !iv_val TYPE clike .
    METHODS _tag_add
      IMPORTING
        !iv_name       TYPE string
        !iv_prefix     TYPE string OPTIONAL
        !iv_attr       TYPE string OPTIONAL
        !iv_flag_close TYPE abap_bool OPTIONAL
      RAISING
        zcx_gl_report_docx .
    METHODS _tag_close
      IMPORTING
        !iv_count TYPE int1 DEFAULT 15 .
    METHODS _write_body
      IMPORTING
        !iv_stamp TYPE clike
      RAISING
        zcx_gl_report_docx .
    METHODS _write_data_to_xml_block
      IMPORTING
        VALUE(iv_tab_pref)  TYPE char255 OPTIONAL
        !iv_sdt_lvl         TYPE numc2 OPTIONAL
        VALUE(io_blok)      LIKE go_blok OPTIONAL
      EXPORTING
        VALUE(ev_blok_name) TYPE char255
      RAISING
        zcx_gl_report_docx .
    METHODS _write_footer
      RAISING
        zcx_gl_report_docx .
    METHODS _write_other
      RAISING
        zcx_gl_report_docx .
    METHODS _write_other_part
      IMPORTING
        !io_part_head TYPE REF TO cl_openxml_part
      RAISING
        zcx_gl_report_docx .
    METHODS ____prepare_xml
      IMPORTING
        !ia_data TYPE any
      RAISING
        zcx_gl_report_docx .
    METHODS ____set_stamp
      IMPORTING
        !iv_stamp TYPE clike
      RAISING
        zcx_gl_report_docx .
ENDCLASS.



CLASS ZCLGL_REPORT_DOCX IMPLEMENTATION.


  METHOD create_hyperlink.
    rv_href = |{ iv_text }{ gc_hyperlink_tab }{ iv_url }|.
  ENDMETHOD.


  METHOD debug.
    IF gv_debug_mode = abap_true.
      BREAK-POINT.                                         "#EC NOBREAK
    ENDIF.
  ENDMETHOD.


  METHOD download.
*--------------------------------------------------------------------*
    go_file->save_on_frontend( iv_string   = gv_final_doc
                               iv_filename = iv_filename ).
*--------------------------------------------------------------------*
    IF iv_open = abap_true.
      go_file->execute( ).
    ENDIF.
*--------------------------------------------------------------------*
  ENDMETHOD.


  METHOD final_check.
*-Проверк сформированного документа--------------------------------------------*
    TRY.
        gv_final_doc = go_doc->get_package_data( ).
      CATCH cx_openxml_format.
        MESSAGE e010 INTO gv_dummy.
        RAISE EXCEPTION TYPE zcx_gl_report_docx
          EXPORTING
            is_sy = sy.
    ENDTRY.
*--------------------------------------------------------------------*
  ENDMETHOD.


  METHOD get_hyperlink.
*--если нет разделителя, то считаем, что текст = ссылке и именного её и переали-------*
    IF iv_href CS gc_hyperlink_tab.
      SPLIT iv_href AT gc_hyperlink_tab INTO ev_text
                                             ev_url.
    ELSE.
      ev_text = iv_href.
      ev_url = iv_href.
    ENDIF.

  ENDMETHOD.


METHOD open_template.
*---Открыйть шаблон и перевести его в XML----------------------------------------*
  TRY.
      CREATE OBJECT go_file.
      go_file->download( iv_template ).
      gv_document = go_file->get_bstring( ).
      go_doc = cl_docx_document=>load_document( iv_data = gv_document ).
      go_documentpart = go_doc->get_maindocumentpart( ).
    CATCH cx_root INTO DATA(lx_err).
      debug( ).
      RAISE EXCEPTION TYPE zcx_gl_report_docx EXPORTING previous = lx_err.
  ENDTRY.
ENDMETHOD.


  METHOD run.
**************************************************************************************************
* I.D. Dochiev, ТНТ, 28.12.2015 17:08:11
* Запуск отчета формата DOC3
* Своя разработка, имеющая следующие достоинства по сравнению с разработкой parazit-а:
*  - Любая вложенность таблиц;
*  - скорость - весь документ формируется в SAP на XML
***************************************************************************************************
    _refresh_global_data( ).
*--------------------------------------------------------------------*
    gth_imgs   = ith_imgs.
    gth_tables = ith_tables.
    gth_values = ith_values.
    gv_debug_mode = iv_debug_mode.
**********************************************************************
*---Открыйть шаблон и перевести его в XML----------------------------------------*
    open_template( iv_template ).
**********************************************************************
    _write_body( iv_stamp = 'iv_stamp' ).
*-Записать данные в подвал документа---------------------------------------------*
*  _write_footer( ).
    _write_other( ).
*--------------------------------------------------------------------*
    final_check( ).
*-Не передали имя файла и закрыть форму, считаем, что файл не нужен.
    ev_content    = gv_final_doc.
*--------------------------------------------------------------------*
    IF iv_open = abap_false.
      RETURN.
    ENDIF.
*-Выгрузка файла на локальную машину или возврат его в EV_CONTENT-----------------*
    download( iv_open = iv_open   " True - форма закрывается после открытия
               iv_filename = iv_filename  ).  " Содержимое файла (если iv_close_form = true)
  ENDMETHOD.


  METHOD _clean.
    DATA: lo_writer     TYPE REF TO cl_sxml_string_writer,
          lv_intrm_part TYPE xstring.
    lo_writer ?= gi_swriter.
    lv_intrm_part = lo_writer->get_output( ).
    CLEAR gv_final_doc.
**--------------------------------------------------------------------*
    IF lv_intrm_part IS NOT INITIAL.
      CALL TRANSFORMATION zgl_doc3_clean
           SOURCE XML lv_intrm_part
           RESULT XML gv_final_doc.
    ENDIF.
**********************************************************************
  ENDMETHOD.


  METHOD _find_variable.
    FIELD-SYMBOLS: <ls_val> LIKE LINE OF gth_values.
    DATA: li_opelem     TYPE REF TO if_sxml_open_element,
          lv_val_key    LIKE <ls_val>-fpath,
          lv_found_t    TYPE abap_bool,
          li_closem     TYPE REF TO if_sxml_close_element,
          lv_sxml_named TYPE c LENGTH 255,
          lv_set_value  TYPE string.
    FIELD-SYMBOLS: <ls_imgs> LIKE LINE OF gth_imgs .

*--------------------------------------------------------------------*
    IF NOT ( ii_sreader IS BOUND AND gi_swriter IS BOUND ).
      RETURN.
    ENDIF.
*--------------------------------------------------------------------*
    lv_val_key = iv_tab_pref.
    READ TABLE gth_values ASSIGNING <ls_val>
                          WITH TABLE KEY fpath = lv_val_key.
    IF sy-subrc = 0.
      READ TABLE gth_imgs ASSIGNING <ls_imgs>
                          WITH TABLE KEY id = <ls_val>-value.
      IF sy-subrc = 0.
        _find_variable_img( is_img     = <ls_imgs>
                            ii_sreader = ii_sreader ).
        RETURN.
      ELSE.
        lv_set_value =  <ls_val>-value.
      ENDIF.
    ELSE.
      lv_set_value = '!!!Нет поля во входных данных'.
    ENDIF.
*--------------------------------------------------------------------*
*  IF iv_tab_pref CS 'HREF'.
*    break dochievid.
*  ENDIF.
*--------------------------------------------------------------------*
    TRY.
        DO.
          gi_snode = ii_sreader->read_next_node( ).
          ii_sreader->current_node( ).
          IF gi_snode IS INITIAL.
            EXIT.
          ENDIF.

          IF gi_snode->type EQ if_sxml_node=>co_nt_element_open.
            li_opelem ?= gi_snode.
            lv_sxml_named = li_opelem->if_sxml_named~qname-name.
            CASE lv_sxml_named.
              WHEN 'hyperlink'.
                _find_variable_hyperlink( CHANGING cv_value = lv_set_value ).
              WHEN 't'.
                _find_variable_prep_val( CHANGING cv_str = lv_set_value ).
                lv_found_t = abap_true.
              WHEN 'document' OR 'body' OR zclgl_report_docx_block_xml=>gc_blok_part.
                CONTINUE.
            ENDCASE.
          ELSEIF gi_snode->type EQ if_sxml_node=>co_nt_element_close.
            li_closem ?= gi_snode.
            lv_sxml_named = li_closem->if_sxml_named~qname-name.
            CASE lv_sxml_named.
              WHEN 'document' OR 'body' OR zclgl_report_docx_block_xml=>gc_blok_part.
                CONTINUE.
            ENDCASE.
          ELSEIF gi_snode->type EQ if_sxml_node=>co_nt_value.
            IF lv_found_t = abap_true.
              gi_svalue_node ?= gi_snode.
              _set_value( lv_set_value ).
              CLEAR: lv_found_t.
              CONTINUE.
            ENDIF.
          ENDIF.
          gi_swriter->write_node( gi_snode ).
        ENDDO.
*--------------------------------------------------------------------*
      CATCH cx_sxml_error.
        debug( ).
        RETURN.
    ENDTRY.

  ENDMETHOD.


  METHOD _find_variable_hyperlink.
    CONSTANTS lc_hyperlink TYPE oxa_opc_relation_type VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'.
    DATA:   lv_text       TYPE string,  " Текст ссылки
            lv_url        TYPE string,
            lv_id         TYPE string,

            ltd_atr       TYPE if_sxml_attribute=>attributes,
            li_attr_named TYPE REF TO if_sxml_named,
            li_opelem     TYPE REF TO if_sxml_open_element.

    FIELD-SYMBOLS: <li_atr>  LIKE LINE OF ltd_atr.
*--------------------------------------------------------------------*
    get_hyperlink( EXPORTING iv_href = cv_value
                   IMPORTING ev_text = lv_text    " Текст ссылки
                             ev_url  = lv_url ).   " Сама ссылка
    cv_value = lv_text.
    lv_id = go_documentpart->add_external_relationship( iv_type      = lc_hyperlink
                                                        iv_targeturi = lv_url ).
    li_opelem ?= gi_snode.
    ltd_atr = li_opelem->get_attributes( ).

    LOOP AT ltd_atr ASSIGNING <li_atr>.
      li_attr_named ?= <li_atr>.
      CHECK <li_atr>->qname-name = 'id'.
      <li_atr>->set_value( lv_id ).
    ENDLOOP.
  ENDMETHOD.


  METHOD _find_variable_img.
    DATA: li_opelem     TYPE REF TO if_sxml_open_element,
          lv_found_t    TYPE abap_bool,
          li_closem     TYPE REF TO if_sxml_close_element,
          lv_sxml_named TYPE c LENGTH 255,
          lv_set_value  TYPE string,
          ltd_atr       TYPE if_sxml_attribute=>attributes,
          li_attr_named TYPE REF TO if_sxml_named,
          lv_xfrm       TYPE abap_bool. ",ls_size TYPE zifgl_reports_img=>gtys_size.
    FIELD-SYMBOLS: <li_atr>  LIKE LINE OF ltd_atr.
*--------------------------------------------------------------------*
    lv_xfrm = abap_false.
*--------------------------------------------------------------------*
    TRY.
        DO.
          CLEAR ltd_atr.
          gi_snode = ii_sreader->read_next_node( ).
          ii_sreader->current_node( ).
          IF gi_snode IS INITIAL.
            EXIT.
          ENDIF.

          IF gi_snode->type EQ if_sxml_node=>co_nt_element_open.
            li_opelem ?= gi_snode.
            lv_sxml_named = li_opelem->if_sxml_named~qname-name.
            CASE lv_sxml_named.
              WHEN 'blip'.
                ltd_atr = li_opelem->get_attributes( ).
                LOOP AT ltd_atr ASSIGNING <li_atr>.
                  li_attr_named ?= <li_atr>.
                  CHECK <li_atr>->qname-name = 'embed'.
                  <li_atr>->set_value( is_img-id_in_doc ).
                ENDLOOP.
*--------------------------------------------------------------------*
              WHEN 'xfrm'.
                lv_xfrm = abap_true.
                "WHEN ' <a:off x="0" y="0"/>
              WHEN 'ext'. "*                 ="1903730" ="1903730"/>
                IF lv_xfrm = abap_true
                  AND is_img-size_in_doc = abap_false
                  AND is_img-height > 0
                  AND is_img-width > 0.
                  ltd_atr = li_opelem->get_attributes( ).
                  LOOP AT ltd_atr ASSIGNING <li_atr>.
                    li_attr_named ?= <li_atr>.
                    CASE <li_atr>->qname-name.
                      WHEN 'cx'.
                        <li_atr>->set_value( |{ is_img-height * 10000 }| ).
                      WHEN 'cy'.
                        <li_atr>->set_value( |{ is_img-width * 10000 }| ).
                    ENDCASE.
                  ENDLOOP.
                ENDIF.
*--------------------------------------------------------------------*
              WHEN 'document' OR 'body' OR zclgl_report_docx_block_xml=>gc_blok_part.
                CONTINUE.
            ENDCASE.
          ELSEIF gi_snode->type EQ if_sxml_node=>co_nt_element_close.
            li_closem ?= gi_snode.
            lv_sxml_named = li_closem->if_sxml_named~qname-name.
            CASE lv_sxml_named.
              WHEN 'document' OR 'body' OR zclgl_report_docx_block_xml=>gc_blok_part.
                CONTINUE.
            ENDCASE.
          ELSEIF gi_snode->type EQ if_sxml_node=>co_nt_value.
            IF lv_found_t = abap_true.
              gi_svalue_node ?= gi_snode.
              _set_value( lv_set_value ).
              CLEAR: lv_found_t.
              CONTINUE.
            ENDIF.
          ENDIF.
          gi_swriter->write_node( gi_snode ).
        ENDDO.
*--------------------------------------------------------------------*
      CATCH cx_sxml_error.
        debug( ).
        RETURN.
    ENDTRY.
  ENDMETHOD.


  METHOD _find_variable_prep_val.
    DATA lv_char TYPE c LENGTH 500.
*--------------------------------------------------------------------*
*  IF sy-uname <> 'DOCHIEVID'.
    RETURN.
*  ENDIF.
*--------------------------------------------------------------------*
    DO 10 TIMES.
      REPLACE cl_abap_char_utilities=>horizontal_tab WITH space INTO cv_str.
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.
*    break dochievid.
*    cv_str = cv_str+1.
      _tag_add( EXPORTING iv_name       = 'tab'
                          iv_flag_close = abap_true ).
    ENDDO.




*  DATA ltd_string TYPE TABLE OF string.
*--------------------------------------------------------------------*
* считаем, что передать табуляцию в конце строки не возможно
*--------------------------------------------------------------------*
*  SPLIT cv_str AT cl_abap_char_utilities=>horizontal_tab
*               INTO TABLE ltd_string.
*  IF sy-subrc <> 0.
*    RETURN.
*  ENDIF.
**--------------------------------------------------------------------*
*  LOOP AT ltd_string INTO cv_str.
*    _tag_add( EXPORTING iv_name       = iv_name
*                        iv_flag_close = abap_true ).
*  ENDLOOP.
*  break dochievid.

  ENDMETHOD.


  METHOD _finish_seq_access.
    DATA: lx_root TYPE REF TO cx_sxml_error.
*--------------------------------------------------------------------*
    IF gi_sreader IS INITIAL.
      RETURN.
    ENDIF.
*--------------------------------------------------------------------*
    DO.
      TRY.
          gi_snode = gi_sreader->read_next_node( ).
          IF gi_snode IS INITIAL.
            EXIT.
          ENDIF.
          go_blok->write_node( gi_snode ).
        CATCH cx_sxml_error INTO lx_root.
          EXIT.
      ENDTRY.
    ENDDO.


  ENDMETHOD.


  METHOD _from_blocks_to_doc.
    DATA: li_opelem        TYPE REF TO if_sxml_open_element,
          li_closem        TYPE REF TO if_sxml_close_element,
          lv_sxml_named    TYPE c LENGTH 255,
          lv_sxml_pref     TYPE c LENGTH 255,
          lts_child        TYPE io_blok->gtyts_child,
          ls_child         LIKE LINE OF lts_child,
          ls_tables        LIKE LINE OF gth_tables,
          lv_curr_tab_name TYPE c LENGTH 255,
          lv_curr_tab_line TYPE i,
          lx_error         TYPE REF TO cx_sxml_error,
          lv_was           TYPE abap_bool.

**********************************************************************
    CHECK io_blok->get_sreader( ) IS BOUND.
    lts_child = io_blok->get_child( ).
*--------------------------------------------------------------------*
    io_blok->get_sreader( abap_true ).
*--------------------------------------------------------------------*
    DO. " бегаем пока не закончиться документ "cx_sxml_error"
      TRY.
          gi_snode = io_blok->get_sreader( )->read_next_node(  ).
          io_blok->get_sreader( )->current_node( ).
          IF gi_snode IS INITIAL.
            EXIT.
          ENDIF.
*--------------------------------------------------------------------*
          IF gi_snode->type EQ if_sxml_node=>co_nt_element_open.
            li_opelem ?= gi_snode.
            lv_sxml_named = li_opelem->if_sxml_named~qname-name.
*--------------------------------------------------------------------*
            IF li_opelem->if_sxml_named~prefix = gc_block_name.
*--------------------------------------------------------------------*
*            li_sxnl_val = li_opelem->get_attribute_value( name  = gc_block_attr_name ).
*            lv_block_name  = li_sxnl_val->get_value( ).
              READ TABLE lts_child INTO ls_child
                                   WITH KEY name = lv_sxml_named
                                   BINARY SEARCH.
              IF sy-subrc = 0.
                IF iv_tab_pref IS INITIAL.
                  lv_curr_tab_name = lv_sxml_named.
                ELSE.
                  lv_curr_tab_name = |{ iv_tab_pref }-{ lv_sxml_named }|.
                ENDIF.

                READ TABLE gth_tables INTO ls_tables
                                      WITH TABLE KEY tpath = lv_curr_tab_name.
                IF sy-subrc = 0.
                  lv_curr_tab_line = 0.
                  DO ls_tables-lines TIMES.
                    ADD 1 TO lv_curr_tab_line.
                    _from_blocks_to_doc( io_blok =  ls_child-o_ref
                                         iv_tab_pref =  |{ lv_curr_tab_name }#{ lv_curr_tab_line }| ).
                  ENDDO.
                ELSE.
*              IF lts_child IS INITIAL. " это конечная ячейка - пишем значние
                  ls_child-o_ref->get_sreader( abap_true )." скинуть буфер, читеть сначала
                  _find_variable( EXPORTING iv_tab_pref = lv_curr_tab_name
                                            ii_sreader = ls_child-o_ref->get_sreader( ) ).
*                                           ii_sreader = io_blok->get_sreader( ) ).

*              ENDIF.
                ENDIF.
              ENDIF.
              CONTINUE.
*          ELSEIF "lv_sxml_named = 'document' OR
*            lv_sxml_named = 'body'.
*            CONTINUE.
            ELSEIF li_opelem->if_sxml_named~prefix = io_blok->gc_blok_part.
              CONTINUE.
            ENDIF.
            IF iv_tab_pref IS INITIAL AND lv_was IS INITIAL.
              _set_attribute_for_first_eleme( li_opelem ).
              lv_was = abap_true.
            ENDIF.
**********************************************************************
          ELSEIF gi_snode->type EQ if_sxml_node=>co_nt_element_close.
            li_closem ?= gi_snode.
            lv_sxml_named = li_closem->if_sxml_named~qname-name.
            lv_sxml_pref = li_closem->if_sxml_named~prefix.
*          IF lv_sxml_named = gc_block_name.
            IF lv_sxml_named = gc_block_name
              OR lv_sxml_named = io_blok->gc_blok_part.
              CONTINUE.
            ELSEIF lv_sxml_pref = gc_block_name
              OR lv_sxml_pref  = io_blok->gc_blok_part.
              CONTINUE.
            ENDIF.
          ENDIF.
*--------------------------------------------------------------------*
          TRY.
              gi_swriter->write_node( gi_snode ).
              IF li_opelem IS BOUND.
                IF li_opelem->if_sxml_named~qname-name = 'body'.
* первым делом поставим штамп/гриф, если надо
*                _set_stamp( gv_stamp ).
                ENDIF.
              ENDIF.
            CATCH cx_sxml_error INTO lx_error.
              debug( ).
              gv_dummy = lx_error->get_text( ).
          ENDTRY.
        CATCH cx_sxml_error INTO lx_error.
          debug( ).
          gv_dummy = lx_error->get_text( ).
*        cl_demo_output=>display_text( lx_error->get_text( ) ).
*        RETURN.
      ENDTRY.

      FREE li_opelem.
    ENDDO.



  ENDMETHOD.


  METHOD _get_nsuri_by_prefix.
    rv_nuri = gi_swriter->get_nsuri_by_prefix( iv_prefix ). "'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
*--------------------------------------------------------------------*
    IF rv_nuri IS INITIAL.
      CASE iv_prefix.
        WHEN 'a'.   rv_nuri = 'http://schemas.openxmlformats.org/drawingml/2006/main'.
        WHEN 'm'.   rv_nuri = 'http://schemas.openxmlformats.org/officeDocument/2006/math'.
        WHEN 'mc'.  rv_nuri = 'http://schemas.openxmlformats.org/markup-compatibility/2006'.
        WHEN 'o'.   rv_nuri = 'urn:schemas-microsoft-com:office:office'.
        WHEN 'pic'. rv_nuri = 'http://schemas.openxmlformats.org/drawingml/2006/picture'.
        WHEN 'r'.   rv_nuri = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'.
        WHEN 'v'.   rv_nuri = 'urn:schemas-microsoft-com:vml'.
        WHEN 'w'.   rv_nuri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'.
        WHEN 'w10'. rv_nuri = 'urn:schemas-microsoft-com:office:word'.
        WHEN 'w14'. rv_nuri = 'http://schemas.microsoft.com/office/word/2010/wordml'.
        WHEN 'wne'. rv_nuri = 'http://schemas.microsoft.com/office/word/2006/wordml'.
        WHEN 'wp'.  rv_nuri = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'.
        WHEN 'wp14'.rv_nuri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'.
        WHEN 'wpc'. rv_nuri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'.
        WHEN 'wpg'. rv_nuri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'.
        WHEN 'wpi'. rv_nuri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk'.
        WHEN 'wps'. rv_nuri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'.
        WHEN OTHERS.
      ENDCASE.
    ENDIF.

  ENDMETHOD.


  METHOD _insert_img.
*--------------------------------------------------------------------*
    DATA: lv_bmp         TYPE zegl_fpm_file_content,
          lx_not_allowed TYPE REF TO cx_openxml_not_allowed,
          lx_format      TYPE REF TO cx_openxml_format,
          lx_not_found   TYPE REF TO cx_openxml_not_found,
          li_log         TYPE REF TO zif_bc_slog.
    FIELD-SYMBOLS: <ls_imgs> LIKE LINE OF gth_imgs.
*--------------------------------------------------------------------*
    IF gth_imgs IS INITIAL.
      RETURN.
    ENDIF.
*--------------------------------------------------------------------*
    li_log = zcl_bc_slog=>get_log( 'ZDOC3_IMG' ).
*--------------------------------------------------------------------*
    LOOP AT gth_imgs ASSIGNING <ls_imgs>.
      TRY.
          <ls_imgs>-o_imagepart = go_documentpart->add_imagepart( cl_oxml_imagepart=>co_content_type_png ).
          <ls_imgs>-o_imagepart->feed_data( <ls_imgs>-bmp ).
          <ls_imgs>-id_in_doc = go_documentpart->get_id_for_part( <ls_imgs>-o_imagepart ).
        CATCH cx_openxml_not_allowed INTO lx_not_allowed.
          li_log->add_msg_text( iv_text            = lx_not_allowed->get_longtext( )
                                iv_msgty           = 'E'
                                iv_split_long_text = abap_true    ).
        CATCH cx_openxml_format INTO lx_format.
          li_log->add_msg_text( iv_text            = lx_format->get_longtext( )
                                iv_msgty           = 'E'
                                iv_split_long_text = abap_true    ).

        CATCH cx_openxml_not_found INTO lx_not_found.
          li_log->add_msg_text( iv_text            = lx_not_found->get_longtext( )
                                iv_msgty           = 'E'
                                iv_split_long_text = abap_true    ).
      ENDTRY.
    ENDLOOP.
*--------------------------------------------------------------------*
    IF li_log->get_error_level( ) = 'E'.
      debug( ).
*    RAISE EXCEPTION TYPE zcx_gl_root EXPORTING gi_log = li_log.
    ENDIF.
*--------------------------------------------------------------------*
  ENDMETHOD.


  METHOD _prep_seq_access.
*--Открываем документ для чтения и создаём документ для записи(выходной документ)*
*--------------------------------------------------------------------*
    DATA: li_open_element TYPE REF TO if_sxml_open_element.
*--------------------------------------------------------------------*
    IF iv_part IS NOT INITIAL.
      gi_sreader ?= cl_sxml_string_reader=>create( iv_part ).
      gi_swriter ?= cl_sxml_string_writer=>create( ).
    ENDIF.
***********************************************************************
**  IF iv_is_document = abap_false.
*    RETURN.
**  ENDIF.
***********************************************************************
** зписываем теги "document" и "body" в начало документа(открываем документ)
*  TRY.
*      li_open_element = gi_swriter->new_open_element( name =  'document' "
*                                                      nsuri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
*                                                      prefix = 'w' ). "
*      li_open_element->set_attribute( EXPORTING name               = 'ignorable'
*                                                nsuri              = 'http://schemas.microsoft.com/office/word/2010/wordml'
*                                                prefix             = 'mc'
*                                                value              = 'w14 wp14' ).
*      gi_swriter->write_node( li_open_element ).
**--------------------------------------------------------------------*
*      li_open_element = gi_swriter->new_open_element( name  = 'body' "
*                                                     nsuri  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
*                                                     prefix = 'w' ). "
*      gi_swriter->write_node( li_open_element ).
*  ENDTRY.


  ENDMETHOD.


  METHOD _refresh_global_data.
    CLEAR:
          gi_snode,
          gi_sreader,
          gi_svalue_node,
          gi_swriter,
          go_blok,
          go_doc,
          go_documentpart,
          go_file,
          gth_tables,
          gth_values,
          gv_debug_mode,
          gv_document,
          gv_final_doc
*        gv_intrm_part,
*        gv_main_part
            .


  ENDMETHOD.


  METHOD _set_attribute_for_first_eleme.
*--------------------------------------------------------------------*
    DEFINE m_xmlns.
      II_ELEMENT->set_attribute( EXPORTING name               = &1
                                           nsuri              = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
                                           prefix             = 'xmlns'
                                           value              = &2 ).
    END-OF-DEFINITION.

*--------------------------------------------------------------------*

    m_xmlns:
              'wps'  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
              'wne'  'http://schemas.microsoft.com/office/word/2006/wordml',
              'wpi'  'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
              'wpg'  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
              'w14'  'http://schemas.microsoft.com/office/word/2010/wordml',
              'w15'  'http://schemas.microsoft.com/office/word/2012/wordml',
              'w'    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
              'w10'  'urn:schemas-microsoft-com:office:word',
              'wp'   'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
              'wp14' 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
              'v'    'urn:schemas-microsoft-com:vml',
              'm'    'http://schemas.openxmlformats.org/officeDocument/2006/math',
              'r'    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'o'    'urn:schemas-microsoft-com:office:office',
              'mc'   'http://schemas.openxmlformats.org/markup-compatibility/2006',
              'wpc'  'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
              'mc'   'http://schemas.openxmlformats.org/markup-compatibility/2006'

.
  ENDMETHOD.


  METHOD _set_blok.
    DATA: lv_value        TYPE string,
          li_open_element TYPE REF TO if_sxml_open_element,
          li_writer       TYPE REF TO if_sxml_writer,
          li_value        TYPE REF TO if_sxml_value_node,
          lx_error        TYPE REF TO cx_sxml_error.
*--------------------------------------------------------------------*
    lv_value = iv_name.
*  DEMO_SXML_OO_WRITER
    TRY.
        li_writer = io_blok->get_swriter( ).
        li_open_element = li_writer->new_open_element( name =   lv_value " gc_block_attr_name
                                                       nsuri = 'http://www.sap.com/abapdemos'
                                                       prefix = gc_block_name ). "
        li_writer->write_node( li_open_element ).
        li_writer->write_node( li_writer->new_close_element( ) ).
      CATCH cx_sxml_error INTO lx_error.
        DATA(lv_error) = lx_error->get_text( ).
        debug( ).
*      cl_demo_output=>display_text( lx_error->get_text( ) ).
*      BREAK-POINT.
*      RETURN.
    ENDTRY.

  ENDMETHOD.


  METHOD _set_value.
    DATA lv_string TYPE string.
*    DATA: lx_root TYPE REF TO cx_sxml_error.
*    DATA: l_xstring TYPE xstring.
*  L_XSTRING = CL_ABAP_CODEPAGE=>CONVERT_TO( I_VAL ).
    lv_string = iv_val.
    IF gi_svalue_node IS BOUND AND gi_swriter IS BOUND.
      TRY.
          gi_svalue_node->if_sxml_value~set_value( lv_string ).
          gi_swriter->write_node( gi_snode ).
        CATCH cx_sxml_error." INTO lx_root.
          EXIT.
      ENDTRY.
    ENDIF.
  ENDMETHOD.


  METHOD _tag_add.
*  _set_stamp_xml( ).
    DATA: li_open_element TYPE REF TO if_sxml_open_element,
          ltd_attr        TYPE TABLE OF string,
          lv_name         TYPE string,
          lv_nsuri        TYPE string,
          lv_prefix       TYPE string,
          lv_value        TYPE string,
          lv_strlen       TYPE i,
          lx_name         TYPE REF TO cx_sxml_name_error,
          lx_state        TYPE REF TO cx_sxml_state_error.

    FIELD-SYMBOLS   <lv_attr> LIKE LINE OF ltd_attr.

*--------------------------------------------------------------------*
    TRY.
        li_open_element = gi_swriter->new_open_element( name =  iv_name "
                                                        nsuri = _get_nsuri_by_prefix( iv_prefix )
                                                        prefix = iv_prefix ). "
*--------------------------------------------------------------------*
        IF iv_attr <> space.
          SPLIT iv_attr AT space INTO TABLE ltd_attr.
          LOOP AT ltd_attr ASSIGNING <lv_attr>.
            CLEAR: lv_name,
                   lv_nsuri,
                   lv_prefix,
                   lv_value.
*  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
*  distT="0" distB="0" distL="0" distR="0"
*  uri="http://schemas.openxmlformats.org/drawingml/2006/picture"'
            FIND REGEX '^([0-z]{1,10}):' IN <lv_attr>.
            IF sy-subrc = 0." IF <ls_attr> CS ':'.
              SPLIT <lv_attr> AT ':' INTO lv_prefix <lv_attr>.
            ENDIF.


            IF <lv_attr> CS '='.
              SPLIT <lv_attr> AT '=' INTO lv_name <lv_attr>.
              lv_strlen = strlen( <lv_attr> ) - 2.
              lv_value = <lv_attr>+1(lv_strlen).
            ELSE.
              lv_name = <lv_attr>.
            ENDIF.
            li_open_element->set_attribute( EXPORTING name   = lv_name
                                                      nsuri  = _get_nsuri_by_prefix( lv_prefix )
                                                      prefix = lv_prefix
                                                      value  = lv_value ).
          ENDLOOP.
        ENDIF.
*--------------------------------------------------------------------*
        gi_swriter->write_node( li_open_element ).
      CATCH cx_sxml_name_error INTO lx_name  .
        debug( ).
        RAISE EXCEPTION TYPE zcx_gl_report_docx
          EXPORTING
            previous = lx_name.
      CATCH cx_sxml_state_error INTO lx_state.
        debug( ).
        RAISE EXCEPTION TYPE zcx_gl_report_docx
          EXPORTING
            previous = lx_state.
    ENDTRY.




    IF iv_flag_close = abap_true.
      _tag_close( 1 ).
    ENDIF.

  ENDMETHOD.


  METHOD _tag_close.
*--------------------------------------------------------------------*
    DO iv_count TIMES.
      TRY.
          gi_swriter->write_node( gi_swriter->new_close_element( ) ).
        CATCH cx_sxml_state_error.
          RETURN.
      ENDTRY.
    ENDDO.

  ENDMETHOD.


  METHOD _write_body.

    DATA: lv_main_part  TYPE xstring.

    lv_main_part = go_documentpart->get_data( ).
*--Открываем документ для чтения и создаём документ для записи(выходной документ)*
    _prep_seq_access( lv_main_part )." iv_is_document = abap_true ).
*--------------------------------------------------------------------*
    _insert_img( ).
*-Разбиваем документ на блоки----------------------------------------------------*
    _write_data_to_xml_block( ).
*--если что-то не перенесли из документа в главный блок, переносим тут-----------*
    _finish_seq_access( ).
**********************************************************************
** первым делом поставим штамп/гриф, если надо
*  _set_stamp( gv_stamp ).
*--Формируем выходной документ на основе блоков XML------------------------------*
    _from_blocks_to_doc( go_blok ).
*--Закроем все открытые теги в вызодном документе--------------------------------*
    _tag_close( ).
**********************************************************************
*-трансформация к DOCX------------------------------------------------------------*
    _clean(  ).
    go_documentpart->feed_data( gv_final_doc ).
  ENDMETHOD.


  METHOD _write_data_to_xml_block.
*--------------------------------------------------------------------*
    DATA: lx_root       TYPE REF TO cx_sxml_error,
          li_opelem     TYPE REF TO if_sxml_open_element,
          li_closem     TYPE REF TO if_sxml_close_element,
          ltd_attr      TYPE if_sxml_attribute=>attributes,
          lv_sxml_named TYPE c LENGTH 255,
          lv_sdt_lvl    LIKE iv_sdt_lvl,
          lo_blok       LIKE io_blok,
          lv_sdtcontent TYPE abap_bool,
          lv_blok_name  LIKE ev_blok_name,
          ltd_at        TYPE if_sxml_attribute=>attributes.

    FIELD-SYMBOLS: <li_attr>  LIKE LINE OF ltd_attr.
**********************************************************************
    lv_sdt_lvl = iv_sdt_lvl + 1.
    IF io_blok IS INITIAL.
* блок не передан, значит первый запуск - его нужно создать
      CREATE OBJECT go_blok EXPORTING iv_name = space.
      lo_blok = go_blok.
      lv_sdtcontent = abap_true.
    ENDIF.
**********************************************************************
*SYSTEM-CALL OBJMGR CLONE go_swriter TO lo_swriter_blok.
    CHECK gi_sreader IS BOUND.
*--------------------------------------------------------------------*
    DO. " бегаем пока не закончиться документ "cx_sxml_error"
      TRY.
          gi_snode = gi_sreader->read_next_node( ).
          gi_sreader->current_node( ).
          IF gi_snode IS INITIAL.
            EXIT.
          ENDIF.
          IF gi_snode->type EQ if_sxml_node=>co_nt_element_open.
            li_opelem ?= gi_snode.
            ltd_at = li_opelem->get_attributes( ).

            lv_sxml_named = li_opelem->if_sxml_named~qname-name.
*--------------------------------------------------------------------*
            IF lv_sxml_named = 'sdt'.
              _write_data_to_xml_block( EXPORTING iv_tab_pref = iv_tab_pref
                                                  iv_sdt_lvl  = lv_sdt_lvl
                                                  io_blok     = lo_blok
                                        IMPORTING ev_blok_name = lv_blok_name ).
              _set_blok( EXPORTING io_blok = lo_blok    " Блоки XML - куда вставить
                                   iv_name = lv_blok_name ).   " имя блока


              CONTINUE.
            ENDIF.
*--------------------------------------------------------------------*
            IF lv_sdt_lvl > 1.
              IF ev_blok_name IS INITIAL.
                IF li_opelem->if_sxml_named~qname-name  = 'tag'.
                  ltd_attr = li_opelem->get_attributes( ).
                  READ TABLE ltd_attr ASSIGNING <li_attr>
                                      WITH KEY  table_line->qname-name = 'val'.
                  IF sy-subrc = 0.
                    ev_blok_name = <li_attr>->get_value( ).
                    TRANSLATE ev_blok_name TO UPPER CASE.
                    CREATE OBJECT lo_blok
                      EXPORTING
                        iv_name   = ev_blok_name
                        io_parent = io_blok.
                    CONTINUE.
                  ENDIF.
                ELSE.
                  CONTINUE.
                ENDIF.
*--------------------------------------------------------------------*
              ELSE.
                IF li_opelem->if_sxml_named~qname-name = 'sdtContent'.
                  lv_sdtcontent = abap_true.
                  CONTINUE.
                ELSE.
* если это SDT то обязателен блок, иначе не инетресно
                  IF lv_sdtcontent = abap_false
                    AND lv_sdt_lvl > 1.
                    CONTINUE.
                  ENDIF.
                ENDIF.
              ENDIF.
              IF lv_sxml_named(3) = 'sdt'.
                CONTINUE.
              ENDIF.
            ENDIF.
**********************************************************************
          ELSEIF gi_snode->type EQ if_sxml_node=>co_nt_element_close.
            li_closem ?= gi_snode.
            lv_sxml_named = li_closem->if_sxml_named~qname-name.
            CASE lv_sxml_named.
              WHEN  'sdt'.
                CONTINUE.
              WHEN  'sdtContent'.
                RETURN.
            ENDCASE.
* если это SDT то обязателен блок, иначе не инетресно
            IF lv_sdtcontent = abap_false
              AND lv_sdt_lvl > 1.
              CONTINUE.
            ENDIF.
          ENDIF.

          IF lo_blok IS BOUND.



            lo_blok->write_node( gi_snode ).
          ENDIF.
        CATCH cx_sxml_error INTO lx_root.
          RETURN.
      ENDTRY.
    ENDDO.


  ENDMETHOD.


  METHOD _write_footer.
    DATA: lv_i       TYPE i,
          lo_footers TYPE REF TO cl_openxml_partcollection,
          lo_footer  TYPE REF TO cl_openxml_part,
          lv_footer  TYPE xstring.

*--------------------------------------------------------------------*
    TRY.
        lo_footers = go_documentpart->get_footerparts( ).
        DO lo_footers->get_count( ) TIMES.
          lo_footer = lo_footers->get_part( lv_i ).
          lv_footer = lo_footer->get_data( ).
*
          ADD 1 TO lv_i.
*--------------------------------------------------------------------*
*--Открываем документ для чтения и создаём документ для записи(выходной документ)*
          _prep_seq_access( lv_footer ).
*-Разбиваем документ на блоки----------------------------------------------------*
          _write_data_to_xml_block(  ).
          CHECK go_blok->get_child( ) IS NOT INITIAL.
*--если что-то не перенесли из документа в главный блок, переносим тут-----------*
          _finish_seq_access( ).
*--Формируем выходной документ на основе блоков XML------------------------------*
          _from_blocks_to_doc( go_blok ).
*--Закроем все открытые теги в вызодном документе--------------------------------*
          _tag_close( ).
**********************************************************************
*-трансформация к DOCX------------------------------------------------------------*
          _clean(  ).
          lo_footer->feed_data( gv_final_doc ).
*--------------------------------------------------------------------*
        ENDDO.
      CATCH cx_openxml_not_found.    "
      CATCH cx_openxml_format.    "
        RETURN.
    ENDTRY.
  ENDMETHOD.


  METHOD _write_other.
*--------------------------------------------------------------------*
    DATA: lo_parts      TYPE REF TO cl_openxml_partcollection,
          lx_xml        TYPE REF TO cx_openxml_format,
          lv_index      TYPE int4,
          lo_index_part TYPE REF TO cl_openxml_part.
*--------------------------------------------------------------------*
    TRY.
        lo_parts = go_documentpart->get_parts( ).
        DO lo_parts->get_count( ) TIMES.
          lo_index_part = lo_parts->get_part( lv_index ).
          _write_other_part( lo_index_part ).
          ADD 1 TO lv_index.
        ENDDO.
      CATCH cx_openxml_format INTO lx_xml.    "
        RAISE EXCEPTION TYPE zcx_gl_report_docx
          EXPORTING
            previous = lx_xml.
    ENDTRY.
  ENDMETHOD.


  METHOD _write_other_part.
    DATA: lv_part           TYPE xstring.
*--------------------------------------------------------------------*
    lv_part = io_part_head->get_data( ).
*--------------------------------------------------------------------*
*--Открываем документ для чтения и создаём документ для записи(выходной документ)*
    _prep_seq_access( lv_part ).
*-Разбиваем документ на блоки----------------------------------------------------*
    _write_data_to_xml_block(  ).
*--если что-то не перенесли из документа в главный блок, переносим тут-----------*
    _finish_seq_access( ).
*--Если нет данных которые стоит переносить, прервать обработку------------------*
    CHECK go_blok->get_child( ) IS NOT INITIAL.
*--Формируем выходной документ на основе блоков XML------------------------------*
    _from_blocks_to_doc( go_blok ).
*--Закроем все открытые теги в вызодном документе--------------------------------*
    _tag_close( ).
**********************************************************************
*-трансформация к DOCX------------------------------------------------------------*
    _clean(  ).
    io_part_head->feed_data( gv_final_doc ).
  ENDMETHOD.


  METHOD ____prepare_xml.
**---Подготовить данных для ввода занчений-----------------------------*
*    DATA: ltd_values TYPE zigl_rptpi_export_values,
*          ltd_tables TYPE zigl_rptpi_export_tables,
*          ls_val     LIKE LINE OF gth_values,
*          lth_img    TYPE zclgl_reports_img=>gtyth_buffer,
*          ls_img3    LIKE LINE OF gth_imgs.
*    FIELD-SYMBOLS: <ls_val> LIKE LINE OF ltd_values,
*                   <ls_img> LIKE LINE OF lth_img.
**--------------------------------------------------------------------*
*    zclgl_reports=>get_value_table( EXPORTING ia_data    = ia_data
*                                    IMPORTING etd_values = ltd_values[]    " API печатных форм: Таблица значений
*                                              etd_tables = ltd_tables[]    " API печатных форм: Строки таблиц таблиц
*                                              eth_img    = lth_img  ).  " Буфер инстанций - таблица
**--Запомним картинки-------------------------------------------------*
*    LOOP AT lth_img ASSIGNING <ls_img>.
*      CLEAR ls_img3.
*      MOVE-CORRESPONDING <ls_img> TO ls_img3.
*      INSERT ls_img3 INTO TABLE gth_imgs.
*    ENDLOOP.
**--Запомним таблицы--------------------------------------------------*
*    INSERT LINES OF ltd_tables INTO TABLE gth_tables.
**--пустая строка для-------------------------------------------------*
*    APPEND INITIAL LINE TO ltd_values.
**--Запомним значения-------------------------------------------------*
*    LOOP AT ltd_values ASSIGNING <ls_val>.
*      IF <ls_val>-fpath <> ls_val-fpath.
*        IF ls_val IS NOT INITIAL.
*          INSERT ls_val INTO TABLE gth_values.
*        ENDIF.
*        CLEAR ls_val.
*        ls_val-fpath = <ls_val>-fpath.
*        ls_val-value = <ls_val>-value.
*      ELSE.
*        ls_val-value = |{ ls_val-value }{ <ls_val>-value }|.
*      ENDIF.
*    ENDLOOP.
***-------------------------------------------------------------------*
  ENDMETHOD.


  METHOD ____set_stamp.
**--------------------------------------------------------------------*
*  DATA: lo_all_imgparts TYPE REF TO cl_openxml_partcollection,
*        lv_count        TYPE int4,
*        lo_part         TYPE REF TO cl_openxml_part,
*        lo_imagepart    TYPE REF TO cl_oxml_imagepart,
**        ls_stamp        TYPE zsgl_stamp,
*        lv_id_msg        TYPE string,
*        lv_width        TYPE i,
*        lv_height       TYPE i .
**--------------------------------------------------------------------*
*  IF iv_stamp EQ space.
*    RETURN.
*  ENDIF.
** Получаем байты и параметры штампа в формате BMP
**  zclgl_reports=>get_stamp( EXPORTING iv_stamp = iv_stamp
**                                      iv_bukrs = '1500'
**                            IMPORTING es_data  = ls_stamp ).
**  lv_width = ls_stamp-width * 10000.
**  lv_height = ls_stamp-height * 10000.
*
**-Выгружаем его в набор файлов-----------------------------*
*  TRY.
*      lo_imagepart = go_documentpart->add_imagepart( cl_oxml_imagepart=>co_content_type_png ).
*
*      lo_imagepart->feed_data( ls_stamp-img_bmp ).
*      lv_id_msg = go_documentpart->get_id_for_part( lo_imagepart ).
*    CATCH cx_openxml_not_allowed.    "
*    CATCH cx_openxml_format.
*    CATCH cx_openxml_not_found.
*  ENDTRY.
**--------------------------------------------------------------------*
**--------------------------------------------------------------------*
***********************************************************************
*  _tag_add( iv_name   = 'p'
*            iv_prefix = 'w'
*            iv_attr = 'w:rsidR="001742E1" w:rsidRDefault="001D665B" w:rsidP="001D665B"' ).
**--------------------------------------------------------------------*
**<w:pPr>
*  _tag_add( iv_name   = 'pPr'
*            iv_prefix = 'w' ).
** <w:jc w:val="right"/>
*  _tag_add( iv_name   = 'jc'
*            iv_prefix = 'w'
*            iv_attr = 'w:val="right"'
*            iv_flag_close = abap_true ).
*
**            </w:pPr>
*  _tag_close( 1 ).
** <w:bookmarkStart w:id="0" w:name="_GoBack"/>
*  _tag_add( iv_name   = 'bookmarkStart'
*            iv_prefix = 'w'
*            iv_attr = 'w:id="0" w:name="_GoBack"'
*            iv_flag_close = abap_true ).
**--------------------------------------------------------------------*
**            <w:r>
*  _tag_add( iv_name   = 'r'
*            iv_prefix = 'w'    ).
**--------------------------------------------------------------------*
**                <w:rPr>
*  _tag_add( iv_name   = 'rPr'
*            iv_prefix = 'w' ).
**<w:noProof/>
*  _tag_add( iv_name   = 'noProof'
*            iv_prefix = 'w'
*            iv_flag_close = abap_true ).
**<w:lang w:eastAsia="ru-RU"/>
*  _tag_add( iv_name   = 'lang'
*            iv_prefix = 'w'
*            iv_attr = 'w:eastAsia="ru-RU"'
*            iv_flag_close = abap_true ).
**                </w:rPr>
*  _tag_close( 1 ).
**--------------------------------------------------------------------*
*  _tag_add( iv_name   = 'drawing'
*            iv_prefix = 'w'    ).
****--------------------------------------------------------------------*
**** <wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="2E94593C" wp14:editId="2DA36C9B">
*  _tag_add( iv_name   = 'inline'
*            iv_prefix = 'wp'
*            iv_attr = 'distT="0" distB="0" distL="0" distR="0"' )." wp14:anchorId="2E94593C" wp14:editId="2DA36C9B"' ).
**--------------------------------------------------------------------*
** <wp:extent cx="1343025" cy="1447800"/>
*  _tag_add( iv_name   = 'extent'
*            iv_prefix = 'wp'
*            iv_attr = |cx="{ lv_width  }" cy="{ lv_height }"|
*            iv_flag_close = abap_true ).
**--------------------------------------------------------------------*
** <wp:effectExtent l="0" t="0" r="9525" b="0"/>
*  _tag_add( iv_name   = 'effectExtent'
*            iv_prefix = 'wp'
*            iv_attr = 'l="0" t="0" r="9525" b="0"'
*            iv_flag_close = abap_true ).
**-<wp:docPr id="stamp1" " name="stamp1"/>-------------------------------------------------------------------*
*  _tag_add( iv_name   = 'docPr'
*            iv_prefix = 'wp'
*            iv_attr = 'id="1" name="stamp1"'
*            iv_flag_close = abap_true ).
****--------------------------------------------------------------------*
** <wp:cNvGraphicFramePr>
*  _tag_add( iv_name   = 'cNvGraphicFramePr'
*            iv_prefix = 'wp' ).
**<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
*  _tag_add( iv_name   = 'graphicFrameLocks'
*          iv_prefix = 'a'
*          iv_attr = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"'
*          iv_flag_close = abap_true ).
**</wp:cNvGraphicFramePr>
*  _tag_close( 1 ).
****--------------------------------------------------------------------*
**    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
*  _tag_add( iv_name   = 'graphic'
*            iv_prefix = 'a'
*            iv_attr = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' ).
**--------------------------------------------------------------------*
**    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
*  _tag_add( iv_name   = 'graphicData'
*            iv_prefix = 'a'
*            iv_attr = 'uri="http://schemas.openxmlformats.org/drawingml/2006/picture"' ).
**--------------------------------------------------------------------*
**    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
*  _tag_add( iv_name   = 'pic'
*            iv_prefix = 'pic'
*            iv_attr = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"').
**--------------------------------------------------------------------*
**    <pic:nvPicPr>
*  _tag_add( iv_name   = 'nvPicPr'
*            iv_prefix = 'pic' ).
**--------------------------------------------------------------------*
**    <pic:cNvPr id="0"/>
*  _tag_add( iv_name   = 'cNvPr'
*            iv_prefix = 'pic'
*            iv_attr = 'id="0" name=""'
*            iv_flag_close = abap_true ).
**--------------------------------------------------------------------*
**    <pic:cNvPicPr/>
*  _tag_add( iv_name   = 'cNvPicPr'
*            iv_prefix = 'pic'
*            iv_flag_close = abap_true ).
**--------------------------------------------------------------------*
**    </pic:nvPicPr>
*  _tag_close( 1 ).
**--------------------------------------------------------------------*
**    <pic:blipFill>
*  _tag_add( iv_name   = 'blipFill'
*            iv_prefix = 'pic' ).
**--------------------------------------------------------------------*
**    <a:blip r:embed="
*  _tag_add( iv_name   = 'blip'
*            iv_prefix = 'a'
*            iv_attr = |r:embed="{ lv_id_msg }"|
*            iv_flag_close = abap_true ).
**--------------------------------------------------------------------*
**    <a:stretch>
*  _tag_add( iv_name   = 'stretch'
*            iv_prefix = 'a' ).
**--------------------------------------------------------------------*
**    <a:fillRect/>
*  _tag_add( iv_name   = 'fillRect'
*            iv_prefix = 'a'
*            iv_flag_close = abap_true ).
**--------------------------------------------------------------------*
**    </a:stretch>
**    </pic:blipFill>
*  _tag_close( 2 ).
**--------------------------------------------------------------------*
**    <pic:spPr>
*  _tag_add( iv_name   = 'spPr'
*            iv_prefix = 'pic' ).
**--------------------------------------------------------------------*
**    <a:xfrm>
*  _tag_add( iv_name   = 'xfrm'
*            iv_prefix = 'a' ).
**--------------------------------------------------------------------*
**    <a:off x="0" y="0"/>
*  _tag_add( iv_name   = 'off'
*            iv_prefix = 'a'
*            iv_attr = 'x="0" y="0"'
*            iv_flag_close = abap_true ).
**--------------------------------------------------------------------*
**    <a:ext cx="lw_x_string " cy="lw_y_string"/>
*  _tag_add( iv_name   = 'ext'
*            iv_prefix = 'a'
*            iv_attr = |cx="{ lv_width * 100 }" cy="{ lv_height * 100 }"|
*            iv_flag_close = abap_true ).
**--------------------------------------------------------------------*
**    </a:xfrm>
*  _tag_close( 1 ).
**--------------------------------------------------------------------*
**    <a:prstGeom prst="rect">
*  _tag_add( iv_name   = 'prstGeom'
*            iv_prefix = 'a'
*            iv_attr = 'prst="rect"' ).
**--------------------------------------------------------------------*
**    <a:avLst/>
*  _tag_add( iv_name   = 'avLst'
*            iv_prefix = 'a'
*            iv_flag_close = abap_true ).
**--------------------------------------------------------------------*
**    </a:prstGeom>
**    </pic:spPr>
**    </pic:pic>
**    </a:graphicData>
**    </a:graphic>
**    </wp:inline>
**    </w:drawing>
**    </w:r>
*  _tag_close( 8 ).
**  <w:bookmarkEnd w:id="0"/>
*  _tag_add( iv_name   = 'bookmarkEnd'
*            iv_prefix = 'w'
*            iv_attr = 'w:id="0"'
*            iv_flag_close = abap_true ).
**    </w:p>
*  _tag_close( 1 ).

  ENDMETHOD.
ENDCLASS.
