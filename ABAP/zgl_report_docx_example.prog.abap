*&---------------------------------------------------------------------*
*& Report  ZGL_REPORT_DOCX_EXAMPLE
*& I.D. Dochiev, 19.01.2018
*& Examples of print form for format DOCX.
*& Данный формат основан на элементах управления содержимом WORD
*&  и позволяет расширить возможности динамического форматирования
*&  отчетов, а именно:
*&        - Поддерживает вложенные таблицы любого уровня вложенности;
*&        -
*&--------------------------------------------------------------------*
REPORT zgl_report_docx_example.
*--------------------------------------------------------------------*
INCLUDE zgl_report_docx_example_cls.
*--------------------------------------------------------------------*
START-OF-SELECTION.
  PERFORM main.
*--------------------------------------------------------------------*
FORM main.
  DATA: lo_data TYPE REF TO lcl_data,
        lo_docx TYPE REF TO zclgl_report_docx.
**********************************************************************
*--------------------------------------------------------------------*
  CREATE OBJECT lo_data.
  lo_data->get_data( ).
*--------------------------------------------------------------------*
  TRY.
      CREATE OBJECT lo_docx.
      lo_docx->run( iv_template   = 'ZGL_EXAMPLE_DOC3'    " ID шаблона (в SMW0)
                    iv_debug_mode = abap_true             " Режим отладки
                    ith_imgs      = lo_data->gth_imgs     " Image - table
                    ith_tables    = lo_data->gth_tables   " Tables with count of lines - table
                    ith_values    = lo_data->gth_values )." Value - table
*--------------------------------------------------------------------*
    CATCH zcx_gl_report_docx INTO DATA(lx_root).
      MESSAGE lx_root->get_longtext( ) TYPE 'E'.
  ENDTRY.
ENDFORM.                    " MAIN
