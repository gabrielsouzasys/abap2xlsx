*--------------------------------------------------------------------*
* REPORT  ZDEMO_EXCEL20
* Demo for method zcl_excel_worksheet-bind_alv:
* export data from ALV (CL_GUI_ALV_GRID) object to excel
*--------------------------------------------------------------------*
report zdemo_excel20.

*----------------------------------------------------------------------*
*       CLASS lcl_handle_events DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class lcl_handle_events definition.
  public section.
    methods:
      on_user_command for event added_function of cl_salv_events
        importing e_salv_function.
endclass.                    "lcl_handle_events DEFINITION

*----------------------------------------------------------------------*
*       CLASS lcl_handle_events IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class lcl_handle_events implementation.
  method on_user_command.
    perform user_command." using e_salv_function text-i08.
  endmethod.                    "on_user_command
endclass.                     "lcl_handle_events IMPLEMENTATION

*--------------------------------------------------------------------*
* DATA DECLARATION
*--------------------------------------------------------------------*

data: lo_excel      type ref to zcl_excel,
      lo_worksheet  type ref to zcl_excel_worksheet,
      lo_alv        type ref to cl_gui_alv_grid,
      lo_salv       type ref to cl_salv_table,
      gr_events     type ref to lcl_handle_events,
      lr_events     type ref to cl_salv_events_table,
      gt_sbook      type table of sbook,
      gt_listheader type slis_t_listheader,
      wa_listheader like line of gt_listheader.

data: l_path            type string,  " local dir
      lv_workdir        type string,
      lv_file_separator type c.

constants:
      lv_default_file_name type string value '20_BindAlv.xlsx'.
*--------------------------------------------------------------------*
*START-OF-SELECTION
*--------------------------------------------------------------------*

start-of-selection.

* get data
* ------------------------------------------

  select *
      into table gt_sbook[]
      from sbook                                        "#EC CI_NOWHERE
      up to 10 rows.

* Display ALV
* ------------------------------------------

  try.
      cl_salv_table=>factory(
        exporting
          list_display = abap_false
        importing
          r_salv_table = lo_salv
        changing
          t_table      = gt_sbook[] ).
    catch cx_salv_msg .
  endtry.

  try.
      lo_salv->set_screen_status(
        exporting
          report        = sy-repid
          pfstatus      = 'ALV_STATUS'
          set_functions = lo_salv->c_functions_all ).
    catch cx_salv_msg .
  endtry.

  lr_events = lo_salv->get_event( ).
  create object gr_events.
  set handler gr_events->on_user_command for lr_events.

  lo_salv->display( ).


*&---------------------------------------------------------------------*
*&      Form  USER_COMMAND
*&---------------------------------------------------------------------*
*      ALV user command
*--------------------------------------------------------------------*
form user_command .
  if sy-ucomm = 'EXCEL'.

* get save file path
    cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = l_path ).
    cl_gui_cfw=>flush( ).
    cl_gui_frontend_services=>directory_browse(
      exporting initial_folder = l_path
      changing selected_folder = l_path ).

    if l_path is initial.
      cl_gui_frontend_services=>get_sapgui_workdir(
        changing sapworkdir = lv_workdir ).
      l_path = lv_workdir.
    endif.

    cl_gui_frontend_services=>get_file_separator(
      changing file_separator = lv_file_separator ).

    concatenate l_path lv_file_separator lv_default_file_name
                into l_path.

* export file to save file path

    perform export_to_excel.

  endif.
endform.                    " USER_COMMAND

*--------------------------------------------------------------------*
* FORM EXPORT_TO_EXCEL
*--------------------------------------------------------------------*
* This subroutine is principal demo session
*--------------------------------------------------------------------*
form export_to_excel.

* create zcl_excel_worksheet object

  create object lo_excel.
  lo_worksheet = lo_excel->get_active_worksheet( ).

* get ALV object from screen

  call function 'GET_GLOBALS_FROM_SLVC_FULLSCR'
    importing
      e_grid = lo_alv.

* build list header

  wa_listheader-typ = 'H'.
  wa_listheader-info = sy-title.
  append wa_listheader to gt_listheader.

  wa_listheader-typ = 'S'.
  wa_listheader-info =  'Created by: ABAP2XLSX Group'.
  append wa_listheader to gt_listheader.

  wa_listheader-typ = 'A'.
  wa_listheader-info =
      'Project hosting at https://cw.sdn.sap.com/cw/groups/abap2xlsx'.
  append wa_listheader to gt_listheader.

* write to excel using method Bin_ALV

  lo_worksheet->bind_alv_ole2(
    exporting
*      I_DOCUMENT_URL          = SPACE " excel template
*      I_XLS                   = 'X' " create in xls format?
      i_save_path             = l_path
      io_alv                  = lo_alv
    it_listheader           = gt_listheader
      i_top                   = 2
      i_left                  = 1
*      I_COLUMNS_HEADER        = 'X'
*      I_COLUMNS_AUTOFIT       = 'X'
*      I_FORMAT_COL_HEADER     =
*      I_FORMAT_SUBTOTAL       =
*      I_FORMAT_TOTAL          =
    exceptions
      miss_guide              = 1
      ex_transfer_kkblo_error = 2
      fatal_error             = 3
      inv_data_range          = 4
      dim_mismatch_vkey       = 5
      dim_mismatch_sema       = 6
      error_in_sema           = 7
      others                  = 8
          ).
  if sy-subrc <> 0.
    message id sy-msgid type sy-msgty number sy-msgno
               with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  endif.

endform.                    "EXPORT_TO_EXCEL
