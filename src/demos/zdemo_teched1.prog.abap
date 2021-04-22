*&---------------------------------------------------------------------*
*& Report  ZDEMO_TECHED1
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_teched1.

*******************************
*   Data Object declaration   *
*******************************

data: lo_excel        type ref to zcl_excel,
      lo_excel_writer type ref to zif_excel_writer,
      lo_worksheet    type ref to zcl_excel_worksheet.

data: lv_file      type xstring,
      lv_bytecount type i,
      lt_file_tab  type solix_tab.

data: lv_full_path      type string,
      lv_workdir        type string,
      lv_file_separator type c.

constants: lv_default_file_name type string value 'TechEd01.xlsx'.

*******************************
* Selection screen management *
*******************************

parameters: p_path type zexcel_export_dir.

at selection-screen on value-request for p_path.
  lv_workdir = p_path.
  cl_gui_frontend_services=>directory_browse( exporting initial_folder  = lv_workdir
                                              changing  selected_folder = lv_workdir ).
  p_path = lv_workdir.

initialization.
  cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

start-of-selection.

  if p_path is initial.
    p_path = lv_workdir.
  endif.
  cl_gui_frontend_services=>get_file_separator( changing file_separator = lv_file_separator ).
  concatenate p_path lv_file_separator lv_default_file_name into lv_full_path.


*******************************
*    abap2xlsx create XLSX    *
*******************************

  " Create excel instance
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Demo01' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).

  " Create xlsx stream
  create object lo_excel_writer type zcl_excel_writer_2007.
  lv_file = lo_excel_writer->write_file( lo_excel ).

*******************************
*            Output           *
*******************************

  " Convert to binary
  lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
  lv_bytecount = xstrlen( lv_file ).

  " Save the file
  cl_gui_frontend_services=>gui_download( exporting bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           changing data_tab     = lt_file_tab ).
