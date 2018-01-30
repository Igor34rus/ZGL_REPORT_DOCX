class ZCLGL_REPORT_DOCX_FILE_XML definition
  public
  create public .

public section.

  data GV_EXTENSION type STRING .
  data GV_TEMPDIR type STRING .
  data GV_BSTRING type XSTRING .

  methods EXECUTE
    raising
      ZCX_GL_REPORT_DOCX .
  methods DOWNLOAD        " download file from web repository
    importing
      value(IV_TEMPLATE) type CLIKE
    raising
      ZCX_GL_REPORT_DOCX .
  methods GET_BSTRING       " returns xstring representation of file
    returning
      value(RV_STRING) type XSTRING .
  methods GET_TEMP_DIR       " chose file storage location
    importing
      value(IV_FILENAME) type CLIKE
    returning
      value(RV_PATH) type STRING .
  methods SAVE_ON_FRONTEND       " upload file to client
    importing
      value(IV_STRING) type XSTRING
      value(IV_FILENAME) type CLIKE
    raising
      ZCX_GL_REPORT_DOCX .
PROTECTED SECTION.
private section.

  data GV_FPATH type STRING .
  data GV_DUMMY type CHAR256 .
ENDCLASS.



CLASS ZCLGL_REPORT_DOCX_FILE_XML IMPLEMENTATION.


  METHOD download.
*  DATA:
*  lo_mr_api                  TYPE REF TO if_mr_api.
*  lo_mr_api = cl_mime_repository_api=>get_api( ).
*  lo_mr_api->get( EXPORTING i_url    = doc_template_url          "Give the template url in MIME repository here
*                  IMPORTING e_content = ev_doc
*                  EXCEPTIONS OTHERS    = 1  ).
    DATA: ls_key TYPE wwwdatatab.
    DATA: ls_mime TYPE w3mime.
    DATA: ltd_mime TYPE STANDARD TABLE OF w3mime.
    FIELD-SYMBOLS <la_data> TYPE any.

    ls_key-relid = 'MI'.
    ls_key-objid = iv_template.

    CALL FUNCTION 'WWWDATA_IMPORT'
      EXPORTING
        key               = ls_key
      TABLES
        mime              = ltd_mime
      EXCEPTIONS
        wrong_object_type = 1
        import_error      = 2
        OTHERS            = 3.

    IF sy-subrc EQ 0.
      LOOP AT ltd_mime INTO ls_mime.
        ASSIGN ls_mime TO <la_data> CASTING TYPE ('X').
        IF <la_data> IS ASSIGNED.
          CONCATENATE gv_bstring <la_data> INTO gv_bstring IN BYTE MODE.
          UNASSIGN <la_data>.
        ENDIF.
      ENDLOOP.
    ELSE.
      MESSAGE e005 WITH iv_template. "005  Отчет &1/&2 не существует (тр. SMW0)
      RAISE EXCEPTION TYPE zcx_gl_report_docx EXPORTING is_sy = sy.
    ENDIF.

  ENDMETHOD.


METHOD execute.
*--------------------------------------------------------------------*
  IF gv_fpath IS INITIAL.
    MESSAGE e001.
    RAISE EXCEPTION TYPE zcx_gl_report_docx EXPORTING is_sy = sy.
  ENDIF.
*--------------------------------------------------------------------*
  cl_gui_frontend_services=>execute( EXPORTING document = gv_fpath
                                     EXCEPTIONS OTHERS  = 1 ) .
  IF sy-subrc NE 0 .
    RAISE EXCEPTION TYPE zcx_gl_report_docx  EXPORTING is_sy = sy.
  ENDIF.
*--------------------------------------------------------------------*
ENDMETHOD.


  method GET_BSTRING.
    rv_string = gv_bstring.
  endmethod.


  method GET_TEMP_DIR.
  DATA: lv_wtitle TYPE string.
  DATA: lv_name TYPE string.
  DATA: lv_fpath TYPE string.
*-Указали папки = передали полынй путь-------------------------------*
  IF iv_filename CS '\'.
    rv_path = iv_filename.
    RETURN.
  ENDIF.
*--------------------------------------------------------------------*
  lv_wtitle = 'Выбор места расположения'.
  lv_name = iv_filename.
*--------------------------------------------------------------------*
  cl_gui_frontend_services=>file_save_dialog( EXPORTING window_title = lv_wtitle
                                                        default_extension = 'docx'
                                                        file_filter = 'docx'
                                                        default_file_name = lv_name
                                               CHANGING filename = lv_name
                                                        path     = gv_tempdir
                                                        fullpath = lv_fpath ).
*--------------------------------------------------------------------*
  cl_gui_cfw=>flush( ).
  rv_path = lv_fpath.
*--------------------------------------------------------------------*
  endmethod.


  method SAVE_ON_FRONTEND.
  DATA: ltd_file_tab TYPE STANDARD TABLE OF solisti1,
        lv_bytecount TYPE i,
        lv_fpath     TYPE string.
*--------------------------------------------------------------------*
  lv_fpath = get_temp_dir( iv_filename ).
*--------------------------------------------------------------------*
  IF lv_fpath IS INITIAL.
    MESSAGE s002 INTO gv_dummy.
    RAISE EXCEPTION TYPE ZCX_GL_REPORT_DOCX EXPORTING is_sy = sy.
  ELSE.
    CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
      EXPORTING
        buffer        = iv_string
      IMPORTING
        output_length = lv_bytecount
      TABLES
        binary_tab    = ltd_file_tab.
    "Save the file
*--------------------------------------------------------------------*
    cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                      filename     = lv_fpath
                                                      filetype     = 'BIN'
                                             CHANGING data_tab     = ltd_file_tab
                                             EXCEPTIONS OTHERS = 8   ).
*--------------------------------------------------------------------*
    IF sy-subrc NE 0.
      MESSAGE ID sy-msgid
              TYPE sy-msgty
              NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4
              INTO gv_dummy.
      RAISE EXCEPTION TYPE ZCX_GL_REPORT_DOCX  EXPORTING is_sy = sy.
    ENDIF.
*--------------------------------------------------------------------*
    gv_fpath = lv_fpath.
  ENDIF.

  endmethod.
ENDCLASS.
