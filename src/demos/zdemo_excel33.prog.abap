*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel33.

type-pools: abap.

data: lo_excel      type ref to zcl_excel,
      lo_worksheet  type ref to zcl_excel_worksheet,
      lo_converter  type ref to zcl_excel_converter,
      lo_autofilter type ref to zcl_excel_autofilter.

data lt_test type table of t005t.

data: l_cell_value type zexcel_cell_value,
      ls_area      type zexcel_s_autofilter_area.

constants: c_airlines type string value 'Airlines'.

constants: gc_save_file_name type string value '33_autofilter.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.

  " Creates active sheet
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Internal table' ).

  select * up to 2 rows from t005t into table lt_test.  "#EC CI_NOWHERE

  create object lo_converter.

  lo_converter->convert( exporting
                            it_table     = lt_test
                            i_row_int    = 1
                            i_column_int = 1
                            io_worksheet = lo_worksheet
                         changing
                            co_excel     = lo_excel ) .

  lo_autofilter = lo_excel->add_new_autofilter( io_sheet = lo_worksheet ) .

  ls_area-row_start = 1.
  ls_area-col_start = 1.
  ls_area-row_end = lo_worksheet->get_highest_row( ).
  ls_area-col_end = lo_worksheet->get_highest_column( ).

  lo_autofilter->set_filter_area( is_area = ls_area ).

  lo_worksheet->get_cell( exporting
                             ip_column    = 'C'
                             ip_row       = 2
                          importing
                             ep_value     = l_cell_value ).
  lo_autofilter->set_value( i_column = 3
                            i_value  = l_cell_value ).


*** Create output
  lcl_output=>output( lo_excel ).
