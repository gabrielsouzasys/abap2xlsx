*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL5
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel5.

data: lo_excel      type ref to zcl_excel,
      lo_worksheet  type ref to zcl_excel_worksheet,
      lo_style_cond type ref to zcl_excel_style_cond.

data: ls_iconset              type zexcel_conditional_iconset.



constants: gc_save_file_name type string value '05_Conditional.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.

  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).

  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.


  ls_iconset-iconset                  = zcl_excel_style_cond=>c_iconset_3trafficlights2.
  ls_iconset-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset-cfvo1_value              = '0'.
  ls_iconset-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset-cfvo2_value              = '33'.
  ls_iconset-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset-cfvo3_value              = '66'.
  ls_iconset-showvalue                = zcl_excel_style_cond=>c_showvalue_true.

  lo_style_cond->mode_iconset  = ls_iconset.
  lo_style_cond->set_range( ip_start_column  = 'C'
                                   ip_start_row     = 4
                                   ip_stop_column   = 'C'
                                   ip_stop_row      = 8 ).


  lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 100 ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'C' ip_value = 1000 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'C' ip_value = 150 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 500 ).


  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule          = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->priority      = 1.
  ls_iconset-iconset           = zcl_excel_style_cond=>c_iconset_3trafficlights2.
  ls_iconset-showvalue         = zcl_excel_style_cond=>c_showvalue_false.
  lo_style_cond->mode_iconset  = ls_iconset.
  lo_style_cond->set_range( ip_start_column  = 'E'
                            ip_start_row     = 4
                            ip_stop_column   = 'E'
                            ip_stop_row      = 8 ).


  lo_worksheet->set_cell( ip_row = 4 ip_column = 'E' ip_value = 100 ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'E' ip_value = 1000 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'E' ip_value = 150 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'E' ip_value = 10 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'E' ip_value = 500 ).



*** Create output
  lcl_output=>output( lo_excel ).
