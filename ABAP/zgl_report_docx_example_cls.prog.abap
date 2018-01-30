*&---------------------------------------------------------------------*
*&  Include           ZGL_REPORT_DOCX_EXAMPLE_CLS
*&---------------------------------------------------------------------*
CLASS lcl_data DEFINITION.
  PUBLIC SECTION.

*--------------------------------------------------------------------*
    DATA: gth_imgs   TYPE zclgl_report_docx=>gtyth_imgs,
          gth_tables TYPE zclgl_report_docx=>gtyth_tables,
          gth_values TYPE zclgl_report_docx=>gtyth_values.
*--------------------------------------------------------------------*
    METHODS: get_data.

  PROTECTED SECTION.
    METHODS:
      get_data_table,
      get_data_sub_table,
      get_data_value,
      get_data_img,
      write_value IMPORTING iv_fpath TYPE clike
                            iv_value TYPE clike .
*--------------------------------------------------------------------*
  PRIVATE SECTION.
    TYPES: BEGIN OF gtys_data_tab.
            INCLUDE TYPE scarr.
    TYPES: td_spfli TYPE TABLE OF spfli WITH DEFAULT KEY,
           END OF gtys_data_tab,
*---gtys_data_tab>
           BEGIN OF gtys_data,
             text      TYPE text255,
             text_down TYPE text255,
             text_top  TYPE text255,
             href      TYPE text255,
             img1      TYPE xstring,
             td_scarr  TYPE TABLE OF gtys_data_tab WITH DEFAULT KEY,
           END OF gtys_data.
*---
    DATA: gs_data  TYPE gtys_data.
ENDCLASS.

CLASS lcl_data IMPLEMENTATION.
  METHOD get_data.
*--------------------------------------------------------------------*
* Prepare data for gtys_data. It's initial structure for comparing data and structure
    DATA ls_data TYPE gtys_data.
*--------------------------------------------------------------------*
    get_data_table( ).
    get_data_value( ).
    get_data_img( ).
*--------------------------------------------------------------------*
  ENDMETHOD.
  METHOD write_value.
    DATA ls_val LIKE LINE OF gth_values.
**********************************************************************
*     Flat value
**********************************************************************
    ls_val-fpath = iv_fpath.
    ls_val-value = iv_value.
    INSERT ls_val INTO TABLE gth_values.

  ENDMETHOD.
  METHOD get_data_table.
    DATA: ls_tab LIKE LINE OF gth_tables.
**********************************************************************
*    Sub Table
**********************************************************************
    ls_tab-tpath = 'TD_SCARR'. "name of SubTable
    ls_tab-lines = 1." lines in SubTable
    INSERT ls_tab INTO TABLE gth_tables.
**********************************************************************
**** Sub Table value
**********************************************************************
*-SubTable-------------------------------------------------------------------*
*---SubTable#
*----line in SubTable-
*-----Field name in SubTable
    write_value( iv_fpath = 'TD_SCARR#1-CARRID'
                 iv_value = 'AA' ).
    write_value( iv_fpath = 'TD_SCARR#1-CARRNAME'
                 iv_value = 'American Airlines' ).
    write_value( iv_fpath = 'TD_SCARR#1-CURRCODE'
                 iv_value = 'USD' ).
*--------------------------------------------------------------------*
    get_data_sub_table( ).
*--------------------------------------------------------------------*
  ENDMETHOD.

  METHOD get_data_sub_table.
    DATA: ls_tab LIKE LINE OF gth_tables.
*-SubSubTable--------------------------------------------------------*
*----SubTable#
*------line in SubTable-
*--------name of SubSubTable
    ls_tab-tpath = 'TD_SCARR#1-TD_SPFLI'.
    ls_tab-lines = 	1."
    INSERT ls_tab INTO TABLE gth_tables.
*--------------------------------------------------------------------*
*-SubSubTable-------------------------------------------------------------------*
*---SubTable#
*----line in SubTable-
*------SubSubTable#
*--------line in SubSubTable-
*-----------Field name in SubTable
    write_value( iv_fpath = 'TD_SCARR#1-CARRID'
                 iv_value = 'AA' )."
    write_value( iv_fpath = 'TD_SCARR#1-CARRNAME'
                 iv_value = 'American Airlines' )."
    write_value( iv_fpath = 'TD_SCARR#1-CURRCODE'
                 iv_value = 'USD' )."
    write_value( iv_fpath = 'TD_SCARR#1-MANDT'
                 iv_value = '100' )."
    write_value( iv_fpath = ''
                 iv_value = '' )."
    write_value( iv_fpath = ''
                 iv_value = '' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-AIRPFROM'
                 iv_value = 'JFK' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-AIRPTO'
                 iv_value = 'SFO' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-ARRTIME'
                 iv_value = '0,584027777777778' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-CARRID'
                 iv_value = 'AA' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-CITYFROM'
                 iv_value = 'NEW YORK' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-CITYTO'
                 iv_value = 'SAN FRANCISCO' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-CONNID'
                 iv_value = '17' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-COUNTRYFR'
                 iv_value = 'US' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-COUNTRYTO'
                 iv_value = 'US' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-DEPTIME'
                 iv_value = '0,458333333333333' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-DISTANCE'
                 iv_value = '2.572' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-DISTID'
                 iv_value = 'МИЛ' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-FLTIME'
                 iv_value = '0,250694444444444' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-FLTYPE'
                 iv_value = '' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-MANDT'
                 iv_value = '100' )."
    write_value( iv_fpath = 'TD_SCARR#1-TD_SPFLI#1-PERIOD'
                 iv_value = '0' )."
    write_value( iv_fpath = 'TD_SCARR#1-URL'
                 iv_value = 'http://www.aa.com' )."
  ENDMETHOD.

  METHOD get_data_img.
    DATA: ls_img LIKE LINE OF gth_imgs,
          ls_val LIKE LINE OF gth_values.
*--------------------------------------------------------------------*
    ls_img-id  = 'IMGID'.
    ls_img-bmp = ''.
    INSERT ls_img INTO TABLE gth_imgs.
**********************************************************************
*     Flat value
**********************************************************************
    ls_val-fpath = 'IMG1'.
    ls_val-value = 'ZGL_EXAMPLE_DOC3_IMG'.
    INSERT ls_val INTO TABLE gth_values.

  ENDMETHOD.

  METHOD get_data_value.
    DATA ls_val LIKE LINE OF gth_values.
**********************************************************************
*     Flat value
**********************************************************************
    write_value( iv_fpath = 'IMG1'
                 iv_value = 'ZGL_EXAMPLE_DOC3_IMG' ).
    write_value( iv_fpath = 'IMG1'
                 iv_value = 'ZGL_EXAMPLE_DOC3_IMG' ).
    write_value( iv_fpath = 'TEXT'
                 iv_value = sy-uname ).
    write_value( iv_fpath = 'TEXT_DOWN'
                 iv_value = 'ABAP Нижний колонтитул' ).
    write_value( iv_fpath = 'TEXT_TOP'
                 iv_value = 'ABAP Верхний колонтитул' ).
    write_value( iv_fpath = 'HREF'
                 iv_value = zclgl_report_docx=>create_hyperlink( iv_text = 'Текст ссылки из DOC3'         " Текст ссылки
                                                                 iv_url  = 'http://sapboard.ru/forum/' ) ). " Сама ссылка
  ENDMETHOD.
ENDCLASS.
