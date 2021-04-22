*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL14
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel14.

data: lo_excel                 type ref to zcl_excel,
      lo_worksheet             type ref to zcl_excel_worksheet,
      lo_style_center          type ref to zcl_excel_style,
      lo_style_right           type ref to zcl_excel_style,
      lo_style_left            type ref to zcl_excel_style,
      lo_style_general         type ref to zcl_excel_style,
      lo_style_bottom          type ref to zcl_excel_style,
      lo_style_middle          type ref to zcl_excel_style,
      lo_style_top             type ref to zcl_excel_style,
      lo_style_justify         type ref to zcl_excel_style,
      lo_style_mixed           type ref to zcl_excel_style,
      lo_style_mixed_wrap      type ref to zcl_excel_style,
      lo_style_rotated         type ref to zcl_excel_style,
      lo_style_shrink          type ref to zcl_excel_style,
      lo_style_indent          type ref to zcl_excel_style,
      lv_style_center_guid     type zexcel_cell_style,
      lv_style_right_guid      type zexcel_cell_style,
      lv_style_left_guid       type zexcel_cell_style,
      lv_style_general_guid    type zexcel_cell_style,
      lv_style_bottom_guid     type zexcel_cell_style,
      lv_style_middle_guid     type zexcel_cell_style,
      lv_style_top_guid        type zexcel_cell_style,
      lv_style_justify_guid    type zexcel_cell_style,
      lv_style_mixed_guid      type zexcel_cell_style,
      lv_style_mixed_wrap_guid type zexcel_cell_style,
      lv_style_rotated_guid    type zexcel_cell_style,
      lv_style_shrink_guid     type zexcel_cell_style,
      lv_style_indent_guid     type zexcel_cell_style.

data: lo_row        type ref to zcl_excel_row.

constants: gc_save_file_name type string value '14_Alignment.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.

  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'sheet1' ).

  "Center
  lo_style_center = lo_excel->add_new_style( ).
  lo_style_center->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
  lv_style_center_guid = lo_style_center->get_guid( ).
  "Right
  lo_style_right = lo_excel->add_new_style( ).
  lo_style_right->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_right.
  lv_style_right_guid = lo_style_right->get_guid( ).
  "Left
  lo_style_left = lo_excel->add_new_style( ).
  lo_style_left->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_left.
  lv_style_left_guid = lo_style_left->get_guid( ).
  "General
  lo_style_general = lo_excel->add_new_style( ).
  lo_style_general->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_general.
  lv_style_general_guid = lo_style_general->get_guid( ).
  "Bottom
  lo_style_bottom = lo_excel->add_new_style( ).
  lo_style_bottom->alignment->vertical = zcl_excel_style_alignment=>c_vertical_bottom.
  lv_style_bottom_guid = lo_style_bottom->get_guid( ).
  "Middle
  lo_style_middle = lo_excel->add_new_style( ).
  lo_style_middle->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
  lv_style_middle_guid = lo_style_middle->get_guid( ).
  "Top
  lo_style_top = lo_excel->add_new_style( ).
  lo_style_top->alignment->vertical = zcl_excel_style_alignment=>c_vertical_top.
  lv_style_top_guid = lo_style_top->get_guid( ).
  "Justify
  lo_style_justify = lo_excel->add_new_style( ).
  lo_style_justify->alignment->vertical = zcl_excel_style_alignment=>c_vertical_justify.
  lv_style_justify_guid = lo_style_justify->get_guid( ).

  "Shrink
  lo_style_shrink = lo_excel->add_new_style( ).
  lo_style_shrink->alignment->shrinktofit = abap_true.
  lv_style_shrink_guid = lo_style_shrink->get_guid( ).

  "Indent
  lo_style_indent = lo_excel->add_new_style( ).
  lo_style_indent->alignment->indent = 5.
  lv_style_indent_guid = lo_style_indent->get_guid( ).

  "Middle / Centered / Wrap
  lo_style_mixed_wrap = lo_excel->add_new_style( ).
  lo_style_mixed_wrap->alignment->horizontal   = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_mixed_wrap->alignment->vertical     = zcl_excel_style_alignment=>c_vertical_center.
  lo_style_mixed_wrap->alignment->wraptext     = abap_true.
  lv_style_mixed_wrap_guid = lo_style_mixed_wrap->get_guid( ).

  "Middle / Centered / Wrap
  lo_style_mixed = lo_excel->add_new_style( ).
  lo_style_mixed->alignment->horizontal   = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_mixed->alignment->vertical     = zcl_excel_style_alignment=>c_vertical_center.
  lv_style_mixed_guid = lo_style_mixed->get_guid( ).

  "Center
  lo_style_rotated = lo_excel->add_new_style( ).
  lo_style_rotated->alignment->horizontal   = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_rotated->alignment->vertical     = zcl_excel_style_alignment=>c_vertical_center.
  lo_style_rotated->alignment->textrotation = 165.                        " -75Ã‚Â° == 90Ã‚Â° + 75Ã‚Â°
  lv_style_rotated_guid = lo_style_rotated->get_guid( ).


  " Set row size for first 7 rows to 40
  do 7 times.
    lo_row = lo_worksheet->get_row( sy-index ).
    lo_row->set_row_height( 40 ).
  enddo.

  "Horizontal alignment
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = 'Centered Text' ip_style = lv_style_center_guid ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'B' ip_value = 'Right Text'    ip_style = lv_style_right_guid ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_value = 'Left Text'     ip_style = lv_style_left_guid ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'B' ip_value = 'General Text'  ip_style = lv_style_general_guid ).

  " Shrink & indent
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'F' ip_value = 'Text shrinked' ip_style = lv_style_shrink_guid ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'F' ip_value = 'Text indented' ip_style = lv_style_indent_guid ).

  "Vertical alignment

  lo_worksheet->set_cell( ip_row = 4 ip_column = 'D' ip_value = 'Bottom Text'    ip_style = lv_style_bottom_guid ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'D' ip_value = 'Middle Text'    ip_style = lv_style_middle_guid ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'D' ip_value = 'Top Text'       ip_style = lv_style_top_guid ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'D' ip_value = 'Justify Text'   ip_style = lv_style_justify_guid ).

  " Wrapped
  lo_worksheet->set_cell( ip_row = 10 ip_column = 'B'
                          ip_value = 'This is a wrapped text centered in the middle'
                          ip_style = lv_style_mixed_wrap_guid ).

  " Rotated
  lo_worksheet->set_cell( ip_row = 10 ip_column = 'D'
                          ip_value = 'This is a centered text rotated by -75Ã‚Â°'
                          ip_style = lv_style_rotated_guid ).

  " forced line break
  data: lv_value type string.
  concatenate 'This is a wrapped text centered in the middle' cl_abap_char_utilities=>cr_lf
    'and a manuall line break.' into lv_value.
  lo_worksheet->set_cell( ip_row = 11 ip_column = 'B'
                          ip_value = lv_value
                          ip_style = lv_style_mixed_guid ).

*** Create output
  lcl_output=>output( lo_excel ).
