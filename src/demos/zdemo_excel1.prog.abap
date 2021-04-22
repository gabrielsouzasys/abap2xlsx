*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL1
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel1.


data: lo_excel     type ref to zcl_excel,
      lo_worksheet type ref to zcl_excel_worksheet,
      lo_hyperlink type ref to zcl_excel_hyperlink,
      lo_column    type ref to zcl_excel_column.

constants: gc_save_file_name type string value '01_HelloWorld.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.
  " Creates active sheet
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
*  lo_worksheet->set_title( ip_title = 'Sheet1' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = sy-datum ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = sy-uzeit ).
  lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = 'https://sapmentors.github.io/abap2xlsx' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = 'Click here to visit abap2xlsx homepage' ip_hyperlink = lo_hyperlink ).

  lo_column = lo_worksheet->get_column( ip_column = 'B' ).
  lo_column->set_width( ip_width = 11 ).



*** Create output
  lcl_output=>output( lo_excel ).
