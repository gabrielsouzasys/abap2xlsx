*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL28
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel28.

data: lo_excel        type ref to zcl_excel,
      lo_excel_writer type ref to zif_excel_writer,
      lo_worksheet    type ref to zcl_excel_worksheet,
      lo_column       type ref to zcl_excel_column.

data: lv_file      type xstring,
      lv_bytecount type i,
      lt_file_tab  type solix_tab.

data: lv_full_path      type string,
      lv_workdir        type string,
      lv_file_separator type c.

constants: lv_default_file_name type string value '28_HelloWorld.csv'.

parameters: p_path type string.

at selection-screen on value-request for p_path.

  cl_gui_frontend_services=>directory_browse( exporting initial_folder  = p_path
                                               changing selected_folder = p_path ).

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

  " Creates active sheet
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet1' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = sy-datum ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = sy-uzeit ).

  lo_column = lo_worksheet->get_column( 'B' ).
  lo_column->set_width( 11 ).

  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Sheet2' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'This is the second sheet' ).

  create object lo_excel_writer type zcl_excel_writer_csv.
  zcl_excel_writer_csv=>set_delimiter( ip_value = cl_abap_char_utilities=>horizontal_tab ).
  zcl_excel_writer_csv=>set_enclosure( ip_value = '''' ).
  zcl_excel_writer_csv=>set_endofline( ip_value = cl_abap_char_utilities=>cr_lf ).

  zcl_excel_writer_csv=>set_active_sheet_index( i_active_worksheet = 2 ).
*  zcl_excel_writer_csv=>set_active_sheet_index_by_name(  I_WORKSHEET_NAME = 'Sheet2' ).

  lv_file = lo_excel_writer->write_file( lo_excel ).

  " Convert to binary
  call function 'SCMS_XSTRING_TO_BINARY'
    exporting
      buffer        = lv_file
    importing
      output_length = lv_bytecount
    tables
      binary_tab    = lt_file_tab.
*  " This method is only available on AS ABAP > 6.40
*  lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
*  lv_bytecount = xstrlen( lv_file ).

  " Save the file
  replace first occurrence of '.csv'  in lv_full_path with '_Sheet2.csv'.
  cl_gui_frontend_services=>gui_download( exporting bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           changing data_tab     = lt_file_tab ).

*  zcl_excel_writer_csv=>set_active_sheet_index( i_active_worksheet = 2 ).
  zcl_excel_writer_csv=>set_active_sheet_index_by_name(  i_worksheet_name = 'Sheet1' ).
  lv_file = lo_excel_writer->write_file( lo_excel ).
  replace first occurrence of '_Sheet2.csv'  in lv_full_path with '_Sheet1.csv'.

  " Convert to binary
  call function 'SCMS_XSTRING_TO_BINARY'
    exporting
      buffer        = lv_file
    importing
      output_length = lv_bytecount
    tables
      binary_tab    = lt_file_tab.
*  " This method is only available on AS ABAP > 6.40
*  lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
*  lv_bytecount = xstrlen( lv_file ).

  " Save the file
  cl_gui_frontend_services=>gui_download( exporting bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           changing data_tab     = lt_file_tab ).
