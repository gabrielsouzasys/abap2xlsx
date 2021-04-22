*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL1
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel30.

data: lo_excel     type ref to zcl_excel,
      lo_worksheet type ref to zcl_excel_worksheet,
      lo_hyperlink type ref to zcl_excel_hyperlink,
      lo_column    type ref to zcl_excel_column.


data: lv_value  type string,
      lv_count  type i value 10,
      lv_packed type p length 16 decimals 1 value '1234567890.5'.

constants: lc_typekind_string type abap_typekind value cl_abap_typedescr=>typekind_string,
           lc_typekind_packed type abap_typekind value cl_abap_typedescr=>typekind_packed,
           lc_typekind_num    type abap_typekind value cl_abap_typedescr=>typekind_num,
           lc_typekind_date   type abap_typekind value cl_abap_typedescr=>typekind_date,
           lc_typekind_s_ls   type string value 's_leading_blanks'.

constants: gc_save_file_name type string value '30_CellDataTypes.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.

  " Creates active sheet
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Cell data types' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Number as String'
                          ip_abap_type = lc_typekind_string ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = '11'
                          ip_abap_type = lc_typekind_string ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'String'
                          ip_abap_type = lc_typekind_string ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = ' String with leading spaces'
                          ip_data_type = lc_typekind_s_ls ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = ' Negative Value'
                          ip_abap_type = lc_typekind_string ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Packed'
                          ip_abap_type = lc_typekind_string ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = '50000.01-'
                          ip_abap_type = lc_typekind_packed ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Number with Percentage'
                          ip_abap_type = lc_typekind_string ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 2 ip_value = '0 %'
                          ip_abap_type = lc_typekind_num ).
  lo_worksheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = 'Date'
                          ip_abap_type = lc_typekind_string ).
  lo_worksheet->set_cell( ip_column = 'E' ip_row = 2 ip_value = '20110831'
                          ip_abap_type = lc_typekind_date ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'Positive Value'
                          ip_abap_type = lc_typekind_string ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = '5000.02'
                          ip_abap_type = lc_typekind_packed ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 3 ip_value = '50 %'
                          ip_abap_type = lc_typekind_num ).

  while lv_count <= 15.
    lv_value = lv_count.
    concatenate 'Positive Value with' lv_value 'Digits' into lv_value separated by space.
    lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_count ip_value = lv_value
                            ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_count ip_value = lv_packed
                            ip_abap_type = lc_typekind_packed ).
    concatenate 'Positive Value with' lv_value 'Digits formated as string' into lv_value separated by space.
    lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_count ip_value = lv_value
                            ip_abap_type = lc_typekind_string ).
    lo_worksheet->set_cell( ip_column = 'E' ip_row = lv_count ip_value = lv_packed
                            ip_abap_type = lc_typekind_string ).
    lv_packed = lv_packed * 10.
    lv_count  = lv_count + 1.
  endwhile.

  lo_column = lo_worksheet->get_column( ip_column = 'A' ).
  lo_column->set_auto_size( abap_true ).
  lo_column = lo_worksheet->get_column( ip_column = 'B' ).
  lo_column->set_auto_size( abap_true ).
  lo_column = lo_worksheet->get_column( ip_column = 'C' ).
  lo_column->set_auto_size( abap_true ).
  lo_column = lo_worksheet->get_column( ip_column = 'D' ).
  lo_column->set_auto_size( abap_true ).
  lo_column = lo_worksheet->get_column( ip_column = 'E' ).
  lo_column->set_auto_size( abap_true ).

*** Create output
  lcl_output=>output( lo_excel ).
