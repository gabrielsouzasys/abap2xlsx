*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL10
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel10.

data: lo_excel      type ref to zcl_excel,
      lo_worksheet  type ref to zcl_excel_worksheet,
      lo_style_cond type ref to zcl_excel_style_cond,
      lo_column     type ref to zcl_excel_column.

data: lt_field_catalog  type zexcel_t_fieldcatalog,
      ls_table_settings type zexcel_s_table_settings,
      ls_iconset        type zexcel_conditional_iconset.

constants: gc_save_file_name type string value '10_iTabFieldCatalog.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.

  field-symbols: <fs_field_catalog> type zexcel_s_fieldcatalog.

  " Creates active sheet
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Internal table' ).

  ls_iconset-iconset                  = zcl_excel_style_cond=>c_iconset_5arrows.
  ls_iconset-cfvo1_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset-cfvo1_value              = '0'.
  ls_iconset-cfvo2_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset-cfvo2_value              = '20'.
  ls_iconset-cfvo3_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset-cfvo3_value              = '40'.
  ls_iconset-cfvo4_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset-cfvo4_value              = '60'.
  ls_iconset-cfvo5_type               = zcl_excel_style_cond=>c_cfvo_type_percent.
  ls_iconset-cfvo5_value              = '80'.
  ls_iconset-showvalue                = zcl_excel_style_cond=>c_showvalue_true.

  "cond style
  lo_style_cond = lo_worksheet->add_new_style_cond( ).
  lo_style_cond->rule         = zcl_excel_style_cond=>c_rule_iconset.
  lo_style_cond->mode_iconset = ls_iconset.
  lo_style_cond->priority     = 1.

  data lt_test type table of sflight.
  select * from sflight into table lt_test.             "#EC CI_NOWHERE

  lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = lt_test ).

  loop at lt_field_catalog assigning <fs_field_catalog>.
    case <fs_field_catalog>-fieldname.
      when 'CARRID'.
        <fs_field_catalog>-position   = 3.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-totals_function = zcl_excel_table=>totals_function_count.
      when 'CONNID'.
        <fs_field_catalog>-position   = 4.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-abap_type  = cl_abap_typedescr=>typekind_int.
        "This avoid the excel warning that the number is formatted as a text: abap2xlsx is not able to recognize numc as a number so it formats the number as a text with
        "the related warning. You can force the type and the framework will correctly format the number as a number
      when 'FLDATE'.
        <fs_field_catalog>-position   = 2.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'PRICE'.
        <fs_field_catalog>-position   = 1.
        <fs_field_catalog>-dynpfld    = abap_true.
        <fs_field_catalog>-totals_function = zcl_excel_table=>totals_function_sum.
        <fs_field_catalog>-style_cond = lo_style_cond->get_guid( ).
      when others.
        <fs_field_catalog>-dynpfld = abap_false.
    endcase.
  endloop.

  ls_table_settings-table_style  = zcl_excel_table=>builtinstyle_medium5.

  lo_worksheet->bind_table( ip_table          = lt_test
                            is_table_settings = ls_table_settings
                            it_field_catalog  = lt_field_catalog ).

  lo_column = lo_worksheet->get_column( ip_column = 'D' ). "make date field a bit wider
  lo_column->set_width( ip_width = 13 ).


*** Create output
  lcl_output=>output( lo_excel ).
