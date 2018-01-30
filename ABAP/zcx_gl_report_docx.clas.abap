class ZCX_GL_REPORT_DOCX definition
  public
  inheriting from CX_STATIC_CHECK
  create public .

public section.

  data IS_SY type SYST .

  methods CONSTRUCTOR
    importing
      !TEXTID like TEXTID optional
      !PREVIOUS like PREVIOUS optional
      !IS_SY type SYST optional .
  methods _GET_TEXT
    returning
      value(EV_RESULT) type STRING .

  methods IF_MESSAGE~GET_LONGTEXT
    redefinition .
protected section.
private section.
ENDCLASS.



CLASS ZCX_GL_REPORT_DOCX IMPLEMENTATION.


  method CONSTRUCTOR.
CALL METHOD SUPER->CONSTRUCTOR
EXPORTING
TEXTID = TEXTID
PREVIOUS = PREVIOUS
.
me->IS_SY = IS_SY .
  endmethod.


METHOD if_message~get_longtext.
  result = _get_text( ).
ENDMETHOD.


METHOD _get_text.
*--------------------------------------------------------------------*
  IF is_sy-msgid IS NOT INITIAL.
    MESSAGE ID     is_sy-msgid
            TYPE   is_sy-msgty
            NUMBER is_sy-msgno
            WITH   is_sy-msgv1
                   is_sy-msgv2
                   is_sy-msgv3
                   is_sy-msgv4
            INTO   ev_result.
  ELSE.
    ev_result =  'some error'.
  ENDIF.
*--------------------------------------------------------------------*
  IF previous IS BOUND.
    ev_result  = |{ ev_result } { previous->get_text( ) }|.
  ENDIF.


*--------------------------------------------------------------------*

ENDMETHOD.
ENDCLASS.
