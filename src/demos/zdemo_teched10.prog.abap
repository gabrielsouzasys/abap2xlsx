*&---------------------------------------------------------------------*
*& Report  ZDEMO_TECHED3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_teched10.

*******************************
*   Data Object declaration   *
*******************************

data: lo_excel        type ref to zcl_excel,
      lo_excel_reader type ref to zif_excel_reader,
      lo_worksheet    type ref to zcl_excel_worksheet.

data: lv_full_path      type string,
      lv_workdir        type string,
      lv_file_separator type c.

data: lt_files type filetable,
      ls_file  type file_table,
      lv_rc    type i,
      lv_value type zexcel_cell_value.

constants: gc_save_file_name type string value 'TechEd01.xlsx'.
include zdemo_excel_outputopt_incl.

start-of-selection.

*******************************
*    abap2xlsx read XLSX    *
*******************************
  create object lo_excel_reader type zcl_excel_reader_2007.
  lo_excel = lo_excel_reader->load_file( lv_full_path ).

  lo_excel->set_active_sheet_index( 1 ).
  lo_worksheet = lo_excel->get_active_worksheet( ).

  lo_worksheet->get_cell( exporting ip_column = 'C'
                                    ip_row    = 10
                          importing ep_value  = lv_value ).

  write: 'abap2xlsx total score is ', lv_value.

*** Create output
  lcl_output=>output( lo_excel ).
