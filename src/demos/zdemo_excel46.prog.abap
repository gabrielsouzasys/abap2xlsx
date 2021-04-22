*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL46
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*
report zdemo_excel46.

constants:
  gc_ws_title_validation type zexcel_sheet_title value 'Validation'.

data:
  lo_excel             type ref to zcl_excel,
  lo_worksheet         type ref to zcl_excel_worksheet,
  lo_range             type ref to zcl_excel_range,
  lv_validation_string type string,
  lo_data_validation   type ref to zcl_excel_data_validation,
  lv_row               type zexcel_cell_row.


constants:
  gc_save_file_name type string value '46_ValidationWarning.xlsx'.

include zdemo_excel_outputopt_incl.


start-of-selection.

*** Sheet Validation

* Creates active sheet
  create object lo_excel.

* Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).

* Set sheet name "Validation"
  lo_worksheet->set_title( gc_ws_title_validation ).


* short validations can be entered as string (<254Char)
  lv_validation_string = '"New York, Rio, Tokyo"'.

* create validation object
  lo_data_validation = lo_worksheet->add_new_data_validation( ).

* create new validation from validation string
  lo_data_validation->type           = zcl_excel_data_validation=>c_type_list.
  lo_data_validation->formula1       = lv_validation_string.
  lo_data_validation->cell_row       = 2.
  lo_data_validation->cell_row_to    = 4.
  lo_data_validation->cell_column    = 'A'.
  lo_data_validation->cell_column_to = 'A'.
  lo_data_validation->allowblank     = 'X'.
  lo_data_validation->showdropdown   = 'X'.
  lo_data_validation->prompttitle    = 'Value list available'.
  lo_data_validation->prompt         = 'Please select a value from the value list'.
  lo_data_validation->errorstyle     = zcl_excel_data_validation=>c_style_warning.
  lo_data_validation->errortitle     = 'Warning'.
  lo_data_validation->error          = 'This value does not exist in current value list.'.

* add some fields with validation
  lv_row = 2.
  while lv_row <= 4.
    lo_worksheet->set_cell( ip_row = lv_row ip_column = 'A' ip_value = 'Select' ).
    lv_row = lv_row + 1.
  endwhile.

*** Create output
  lcl_output=>output( lo_excel ).
