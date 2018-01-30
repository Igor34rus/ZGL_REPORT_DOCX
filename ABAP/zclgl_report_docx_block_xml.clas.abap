class ZCLGL_REPORT_DOCX_BLOCK_XML definition
  public
  create public .

public section.

  types:
    BEGIN OF gtys_child,
        name  TYPE char255,
        num   TYPE i,
        o_ref TYPE REF TO zclgl_report_docx_block_xml,
      END OF gtys_child .
  types:
    gtyts_child TYPE SORTED TABLE OF gtys_child WITH UNIQUE KEY name
                                                    WITH  UNIQUE SORTED KEY num COMPONENTS num .

  data GV_NAME type GTYS_CHILD-NAME .
  data GV_COUNT_CHILD type GTYS_CHILD-NUM .
  constants GC_BLOK_PART type STRING value 'BLOK_PART' ##NO_TEXT.

  methods CONSTRUCTOR
    importing
      value(IO_PARENT) type ref to ZCLGL_REPORT_DOCX_BLOCK_XML optional
      value(IV_NAME) type CLIKE
    raising
      ZCX_GL_REPORT_DOCX .
  methods GET_PARENT
    returning
      value(RO_PARENT) type ref to ZCLGL_REPORT_DOCX_BLOCK_XML .
  methods GET_CHILD
    returning
      value(RTS_CHILD) type GTYTS_CHILD .
  methods ADD_CHILD
    importing
      !IO_PARENT type ref to ZCLGL_REPORT_DOCX_BLOCK_XML
    raising
      ZCX_GL_REPORT_DOCX .
  methods GET_SREADER
    importing
      value(IV_FIRST_ROW) type ABAP_BOOL optional
    returning
      value(RI_SREADER) type ref to IF_SXML_READER .
  methods GET_SWRITER
    returning
      value(RI_SWRITER) type ref to IF_SXML_WRITER .
  methods WRITE_NODE
    importing
      !II_NODE type ref to IF_SXML_NODE .
PROTECTED SECTION.
PRIVATE SECTION.

  DATA go_parent TYPE REF TO zclgl_report_docx_block_xml.
  DATA gts_child TYPE gtyts_child .
  DATA gi_sreader TYPE REF TO if_sxml_reader .
  DATA gi_swriter TYPE REF TO if_sxml_writer .
  DATA gv_xml_data TYPE xstring .
ENDCLASS.



CLASS ZCLGL_REPORT_DOCX_BLOCK_XML IMPLEMENTATION.


  method ADD_CHILD.
    DATA ls_child LIKE LINE OF gts_child.
*--------------------------------------------------------------------*
    READ TABLE gts_child TRANSPORTING NO FIELDS
                         WITH KEY name = io_parent->gv_name
                         BINARY SEARCH.
    IF sy-subrc = 0.
      "BREAK-POINT.
*      MESSAGE 'Блок уже описан' TYPE 'E'.
    ENDIF.
*--------------------------------------------------------------------*
    ADD 1 TO gv_count_child.
*--------------------------------------------------------------------*
    ls_child-name = io_parent->gv_name.
    ls_child-num  = gv_count_child.
    ls_child-o_ref = io_parent.
    INSERT ls_child INTO TABLE gts_child.
  endmethod.


  method CONSTRUCTOR.
  gv_name = iv_name.

  IF io_parent IS BOUND.
    go_parent = io_parent.
    io_parent->add_child( me ).
  ENDIF.
  endmethod.


  method GET_CHILD.
  rts_child = gts_child.
  endmethod.


  method GET_PARENT.
  ro_parent = go_parent.
  endmethod.


  method GET_SREADER.
  DATA: lo_writer TYPE REF TO cl_sxml_string_writer.


  IF gi_sreader IS NOT BOUND OR iv_first_row = abap_true.
    TRY.
        gi_swriter->write_node( gi_swriter->new_close_element( ) )." body
      CATCH cx_sxml_state_error.
    ENDTRY.
    TRY.
        gi_swriter->write_node( gi_swriter->new_close_element( ) )." document
      CATCH cx_sxml_state_error.
    ENDTRY.


    lo_writer ?= gi_swriter.
    gv_xml_data = lo_writer->get_output( ).
    gi_sreader ?= cl_sxml_string_reader=>create( gv_xml_data ).
  ENDIF.

  ri_sreader = gi_sreader.
  endmethod.


  method GET_SWRITER.

  DATA: lx_error TYPE REF TO cx_sxml_state_error.

  IF gi_swriter IS NOT BOUND.
    gi_swriter ?= cl_sxml_string_writer=>create( ).
*--------------------------------------------------------------------*
    DATA: li_open_element TYPE REF TO if_sxml_open_element.
    DATA li_value TYPE REF TO if_sxml_value_node.
**  <?xml version="1.0" encoding="UTF-8" standalone="true"?>
    TRY.
*        li_open_element = gi_swriter->new_open_element( name =  'document' "
*                                                       nsuri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
*                                                       prefix = 'w' ). "
*        li_open_element->set_attribute( EXPORTING name               = 'ignorable'
*                                                  nsuri              = 'http://schemas.microsoft.com/office/word/2010/wordml'
*                                                  prefix             = 'mc'
*                                                  value              = 'w14 wp14' ).
*        DEFINE m_xmlns.
*          li_open_element->set_attribute( EXPORTING name               = &1
*                                                    nsuri              = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
*                                                    prefix             = 'xmlns'
*                                                    value              = &2 ).
*        END-OF-DEFINITION.
*
*        m_xmlns:
*                'wps' 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
*                'wne' 'http://schemas.microsoft.com/office/word/2006/wordml',
*                'wpi' 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
*                'wpg' 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
*                'w14' 'http://schemas.microsoft.com/office/word/2010/wordml',
*                'w' 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
*                'w10' 'urn:schemas-microsoft-com:office:word',
*                'wp' 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
*                'wp14' 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
*                'v' 'urn:schemas-microsoft-com:vml',
*                'm' 'http://schemas.openxmlformats.org/officeDocument/2006/math',
*                'r' 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
*                'o' 'urn:schemas-microsoft-com:office:office', 'mc' 'http://schemas.openxmlformats.org/markup-compatibility/2006',
*                'wpc' 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'.
*        gi_swriter->write_node( li_open_element ).
**--------------------------------------------------------------------*
*        li_open_element = gi_swriter->new_open_element( name =   'body' "
*                                                       nsuri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
*                                                       prefix = 'w' ). "
*        gi_swriter->write_node( li_open_element ).
**      li_writer->write_node( li_writer->new_close_element( ) ).
*--------------------------------------------------------------------*

        li_open_element = gi_swriter->new_open_element( name =  gc_blok_part
                                                        nsuri = 'http://www.sap.com/abapdemos'
                                                        prefix = gc_blok_part )." gc_block_name ). "
        gi_swriter->write_node( li_open_element ).
      CATCH  cx_sxml_name_error.    "
      CATCH cx_sxml_state_error INTO lx_error.

        "BREAK-POINT.
        cl_demo_output=>display_text( lx_error->get_text( ) ).
        RETURN.
    ENDTRY.


  ENDIF.




  ri_swriter = gi_swriter.
  endmethod.


  method WRITE_NODE.

  DATA: lx_root  TYPE REF TO cx_sxml_error.

  TRY.
      get_swriter( )->write_node( ii_node ).
    CATCH cx_sxml_error INTO lx_root.
*--------------------------------------------------------------------*
      cl_demo_output=>display_text( lx_root->get_text( ) ).
      RETURN.
  ENDTRY.

*  ##todo_Временна_для_тест_потом_удалить
*
*  DATA: lo_writer TYPE REF TO cl_sxml_string_writer.
*  lo_writer ?= get_swriter( ).
*  gv_xml_data = lo_writer->get_output( ).
*  IF gv_xml_data IS NOT INITIAL.
*    CALL TRANSFORMATION zgl_doc3_clean
*         SOURCE XML gv_xml_data
*         RESULT XML gv_xml_data.
*  ENDIF.
  endmethod.
ENDCLASS.
