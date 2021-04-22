*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel3.

type-pools: abap.

data: lo_excel     type ref to zcl_excel,
      lo_worksheet type ref to zcl_excel_worksheet,
      lo_column    type ref to zcl_excel_column.

data: ls_table_settings       type zexcel_s_table_settings.


data: lv_title type zexcel_sheet_title,
      lt_carr  type table of scarr,
      row      type zexcel_cell_row value 2,
      lo_range type ref to zcl_excel_range.
data: lo_data_validation  type ref to zcl_excel_data_validation.
field-symbols: <carr> like line of lt_carr.

constants: c_airlines type string value 'Airlines'.


constants: gc_save_file_name type string value '03_iTab.xlsx'.
include zdemo_excel_outputopt_incl.

parameters: p_empty type flag.

start-of-selection.
  " Creates active sheet
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Internal table' ).

  data lt_test type table of sflight.

  if p_empty <> abap_true.
    select * from sflight into table lt_test.           "#EC CI_NOWHERE
  endif.

  ls_table_settings-table_style       = zcl_excel_table=>builtinstyle_medium2.
  ls_table_settings-show_row_stripes  = abap_true.
  ls_table_settings-nofilters         = abap_true.

  lo_worksheet->bind_table( ip_table          = lt_test
                            is_table_settings = ls_table_settings ).

  lo_worksheet->freeze_panes( ip_num_rows = 3 ). "freeze column headers when scrolling

  lo_column = lo_worksheet->get_column( ip_column = 'E' ). "make date field a bit wider
  lo_column->set_width( ip_width = 11 ).
  " Add another table for data validations
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lv_title = 'Data Validation'.
  lo_worksheet->set_title( lv_title ).
  lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = c_airlines ).
  select * from scarr into table lt_carr.               "#EC CI_NOWHERE
  loop at lt_carr assigning <carr>.
    lo_worksheet->set_cell( ip_row = row ip_column = 'A' ip_value = <carr>-carrid ).
    row = row + 1.
  endloop.
  row = row - 1.
  lo_range            = lo_excel->add_new_range( ).
  lo_range->name      = c_airlines.
  lo_range->set_value( ip_sheet_name    = lv_title
                       ip_start_column  = 'A'
                       ip_start_row     = 2
                       ip_stop_column   = 'A'
                       ip_stop_row      = row ).
  " Set Data Validation
  lo_excel->set_active_sheet_index( 1 ).
  lo_worksheet = lo_excel->get_active_worksheet( ).

  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
  lo_data_validation->formula1    = c_airlines.
  lo_data_validation->cell_row    = 4.
  lo_data_validation->cell_column = 'C'.

*** Create output
  lcl_output=>output( lo_excel ).
