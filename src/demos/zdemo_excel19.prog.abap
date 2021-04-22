*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL19
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel19.

type-pools: abap.

data: lo_excel     type ref to zcl_excel,
      lo_worksheet type ref to zcl_excel_worksheet.


constants: gc_save_file_name type string value '19_SetActiveSheet.xlsx'.
include zdemo_excel_outputopt_incl.

parameters: p_noout type abap_bool default abap_true.


start-of-selection.

  create object lo_excel.

  " First Worksheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'First' ).
  lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'This is Sheet 1' ).

  " Second Worksheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Second' ).
  lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'This is Sheet 2' ).

  " Third Worksheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Third' ).
  lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'This is Sheet 3' ).

  if p_noout eq abap_false.
    " lo_excel->set_active_sheet_index_by_name( data_sheet_name ).
    data: active_sheet_index type zexcel_active_worksheet.
    active_sheet_index = lo_excel->get_active_sheet_index( ).
    write: 'Sheet Index before: ', active_sheet_index.
  endif.
  lo_excel->set_active_sheet_index( '2' ).
  if p_noout eq abap_false.
    active_sheet_index = lo_excel->get_active_sheet_index( ).
    write: 'Sheet Index after: ', active_sheet_index.
  endif.


*** Create output
  lcl_output=>output( lo_excel ).
