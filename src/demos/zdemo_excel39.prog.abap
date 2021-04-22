*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL16
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel39.

data: lo_excel     type ref to zcl_excel,
      lo_worksheet type ref to zcl_excel_worksheet,
      lo_drawing   type ref to zcl_excel_drawing.

data lv_value type i.

data: ls_io type skwf_io.

data: ls_upper type zexcel_drawing_location,
      ls_lower type zexcel_drawing_location.

data: lo_bar1         type ref to zcl_excel_graph_bars,
      lo_bar1_stacked type ref to zcl_excel_graph_bars,
      lo_bar2         type ref to zcl_excel_graph_bars,
      lo_pie          type ref to zcl_excel_graph_pie,
      lo_line         type ref to zcl_excel_graph_line.

constants: gc_save_file_name type string value '39_Charts.xlsx'.
include zdemo_excel_outputopt_incl.

start-of-selection.

  " Create a pie chart and series
  create object lo_pie.

  call method lo_pie->create_serie
    exporting
      ip_order        = 0
      ip_sheet        = 'Values'
      ip_lbl_from_col = 'B'
      ip_lbl_from_row = '1'
      ip_lbl_to_col   = 'B'
      ip_lbl_to_row   = '3'
      ip_ref_from_col = 'A'
      ip_ref_from_row = '1'
      ip_ref_to_col   = 'A'
      ip_ref_to_row   = '3'
      ip_sername      = 'My serie 1'.

  " Set style
  lo_pie->set_style( zcl_excel_graph=>c_style_15 ).

  " Create a bar chart, series and axes
  create object lo_bar1.

  call method lo_bar1->create_serie
    exporting
      ip_order            = 0
      ip_invertifnegative = zcl_excel_graph_bars=>c_invertifnegative_no
      ip_lbl              = 'Values!$D$1:$D$3'
      ip_ref              = 'Values!$C$1:$C$3'
      ip_sername          = 'My serie 1'.

  call method lo_bar1->create_serie
    exporting
      ip_order            = 1
      ip_invertifnegative = zcl_excel_graph_bars=>c_invertifnegative_no
      ip_lbl              = 'Values!$B$1:$B$3'
      ip_ref              = 'Values!$A$1:$A$3'
      ip_sername          = 'My serie 2'.

  call method lo_bar1->create_ax
    exporting
*     ip_axid =
      ip_type = zcl_excel_graph_bars=>c_catax
*     ip_orientation   =
*     ip_delete        =
*     ip_axpos         =
*     ip_formatcode    =
*     ip_sourcelinked  =
*     ip_majortickmark =
*     ip_minortickmark =
*     ip_ticklblpos    =
*     ip_crossax       =
*     ip_crosses       =
*     ip_auto =
*     ip_lblalgn       =
*     ip_lbloffset     =
*     ip_nomultilvllbl =
*     ip_crossbetween  =
    .

  call method lo_bar1->create_ax
    exporting
*     ip_axid =
      ip_type = zcl_excel_graph_bars=>c_valax
*     ip_orientation   =
*     ip_delete        =
*     ip_axpos         =
*     ip_formatcode    =
*     ip_sourcelinked  =
*     ip_majortickmark =
*     ip_minortickmark =
*     ip_ticklblpos    =
*     ip_crossax       =
*     ip_crosses       =
*     ip_auto =
*     ip_lblalgn       =
*     ip_lbloffset     =
*     ip_nomultilvllbl =
*     ip_crossbetween  =
    .

  " Set style
  lo_bar1->set_style( zcl_excel_graph=>c_style_default ).
  lo_bar1->set_title( ip_value = 'TITLE!' ).

  " Set label to none
  lo_bar1->set_print_lbl( zcl_excel_graph_bars=>c_show_false ).

* Same barchart - but this time stacked
  create object lo_bar1_stacked.

  call method lo_bar1_stacked->create_serie
    exporting
      ip_order            = 0
      ip_invertifnegative = zcl_excel_graph_bars=>c_invertifnegative_no
      ip_lbl              = 'Values!$D$1:$D$3'
      ip_ref              = 'Values!$C$1:$C$3'
      ip_sername          = 'My serie 1'.

  call method lo_bar1_stacked->create_serie
    exporting
      ip_order            = 1
      ip_invertifnegative = zcl_excel_graph_bars=>c_invertifnegative_no
      ip_lbl              = 'Values!$B$1:$B$3'
      ip_ref              = 'Values!$A$1:$A$3'
      ip_sername          = 'My serie 2'.

  call method lo_bar1_stacked->create_ax
    exporting
      ip_type = zcl_excel_graph_bars=>c_catax.

  call method lo_bar1_stacked->create_ax
    exporting
      ip_type = zcl_excel_graph_bars=>c_valax.

  " Set style
  lo_bar1_stacked->set_style( zcl_excel_graph=>c_style_default ).

  " Set label to none
  lo_bar1_stacked->set_print_lbl( zcl_excel_graph_bars=>c_show_false ).

  " Make it stacked
  lo_bar1_stacked->ns_groupingval = zcl_excel_graph_bars=>c_groupingval_stacked.


  " Create a bar chart, series and axes
  create object lo_bar2.

  call method lo_bar2->create_serie
    exporting
      ip_order            = 0
      ip_invertifnegative = zcl_excel_graph_bars=>c_invertifnegative_yes
      ip_lbl              = 'Values!$D$1:$D$3'
      ip_ref              = 'Values!$C$1:$C$3'
      ip_sername          = 'My serie 1'.

  call method lo_bar2->create_ax
    exporting
*     ip_axid =
      ip_type = zcl_excel_graph_bars=>c_catax
*     ip_orientation   =
*     ip_delete        =
*     ip_axpos         =
*     ip_formatcode    =
*     ip_sourcelinked  =
*     ip_majortickmark =
*     ip_minortickmark =
*     ip_ticklblpos    =
*     ip_crossax       =
*     ip_crosses       =
*     ip_auto =
*     ip_lblalgn       =
*     ip_lbloffset     =
*     ip_nomultilvllbl =
*     ip_crossbetween  =
    .

  call method lo_bar2->create_ax
    exporting
*     ip_axid =
      ip_type = zcl_excel_graph_bars=>c_valax
*     ip_orientation   =
*     ip_delete        =
*     ip_axpos         =
*     ip_formatcode    =
*     ip_sourcelinked  =
*     ip_majortickmark =
*     ip_minortickmark =
*     ip_ticklblpos    =
*     ip_crossax       =
*     ip_crosses       =
*     ip_auto =
*     ip_lblalgn       =
*     ip_lbloffset     =
*     ip_nomultilvllbl =
*     ip_crossbetween  =
    .

  " Set layout
  lo_bar2->set_show_legend_key( zcl_excel_graph_bars=>c_show_true ).
  lo_bar2->set_show_values( zcl_excel_graph_bars=>c_show_true ).
  lo_bar2->set_show_cat_name( zcl_excel_graph_bars=>c_show_true ).
  lo_bar2->set_show_ser_name( zcl_excel_graph_bars=>c_show_true ).
  lo_bar2->set_show_percent( zcl_excel_graph_bars=>c_show_true ).
  lo_bar2->set_varycolor( zcl_excel_graph_bars=>c_show_true ).

  " Create a line chart, series and axes
  create object lo_line.

  call method lo_line->create_serie
    exporting
      ip_order   = 0
      ip_symbol  = zcl_excel_graph_line=>c_symbol_auto
      ip_smooth  = zcl_excel_graph_line=>c_show_false
      ip_lbl     = 'Values!$D$1:$D$3'
      ip_ref     = 'Values!$C$1:$C$3'
      ip_sername = 'My serie 1'.

  call method lo_line->create_serie
    exporting
      ip_order   = 1
      ip_symbol  = zcl_excel_graph_line=>c_symbol_none
      ip_smooth  = zcl_excel_graph_line=>c_show_false
      ip_lbl     = 'Values!$B$1:$B$3'
      ip_ref     = 'Values!$A$1:$A$3'
      ip_sername = 'My serie 2'.

  call method lo_line->create_serie
    exporting
      ip_order   = 2
      ip_symbol  = zcl_excel_graph_line=>c_symbol_auto
      ip_smooth  = zcl_excel_graph_line=>c_show_false
      ip_lbl     = 'Values!$F$1:$F$3'
      ip_ref     = 'Values!$E$1:$E$3'
      ip_sername = 'My serie 3'.

  call method lo_line->create_ax
    exporting
*     ip_axid =
      ip_type = zcl_excel_graph_line=>c_catax
*     ip_orientation   =
*     ip_delete        =
*     ip_axpos         =
*     ip_majortickmark =
*     ip_minortickmark =
*     ip_ticklblpos    =
*     ip_crossax       =
*     ip_crosses       =
*     ip_auto =
*     ip_lblalgn       =
*     ip_lbloffset     =
*     ip_nomultilvllbl =
*     ip_crossbetween  =
    .

  call method lo_line->create_ax
    exporting
*     ip_axid =
      ip_type = zcl_excel_graph_line=>c_valax
*     ip_orientation   =
*     ip_delete        =
*     ip_axpos         =
*     ip_formatcode    =
*     ip_sourcelinked  =
*     ip_majortickmark =
*     ip_minortickmark =
*     ip_ticklblpos    =
*     ip_crossax       =
*     ip_crosses       =
*     ip_auto =
*     ip_lblalgn       =
*     ip_lbloffset     =
*     ip_nomultilvllbl =
*     ip_crossbetween  =
    .

  " Creates active sheet
  create object lo_excel.

  " Get active sheet (Pie sheet)
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'PieChart' ).

  " Create global drawing, set type as pie chart, assign chart, set position and media type
  lo_drawing = lo_worksheet->excel->add_new_drawing(
                    ip_type  = zcl_excel_drawing=>type_chart
                    ip_title = 'CHART PIE' ).
  lo_drawing->graph = lo_pie.
  lo_drawing->graph_type = zcl_excel_drawing=>c_graph_pie.

  "Set chart position (anchor 2 cells)
  ls_lower-row = 30.
  ls_lower-col = 20.
  lo_drawing->set_position2(
    exporting
      ip_from   = ls_upper
      ip_to     = ls_lower ).

  lo_drawing->set_media(
    exporting
      ip_media_type = zcl_excel_drawing=>c_media_type_xml ).

  lo_worksheet->add_drawing( lo_drawing ).

  " BarChart1 sheet

  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'BarChart1' ).

  " Create global drawing, set type as bar chart, assign chart, set position and media type
  lo_drawing = lo_worksheet->excel->add_new_drawing(
                    ip_type  = zcl_excel_drawing=>type_chart
                    ip_title = 'CHART BARS WITH 2 SERIES' ).
  lo_drawing->graph = lo_bar1.
  lo_drawing->graph_type = zcl_excel_drawing=>c_graph_bars.

  "Set chart position (anchor 2 cells)
  ls_upper-row = 0.
  ls_upper-col = 11.
  ls_lower-row = 22.
  ls_lower-col = 21.
  lo_drawing->set_position2(
    exporting
      ip_from   = ls_upper
      ip_to     = ls_lower ).

  lo_drawing->set_media(
    exporting
      ip_media_type = zcl_excel_drawing=>c_media_type_xml ).

  lo_worksheet->add_drawing( lo_drawing ).

  lo_drawing = lo_worksheet->excel->add_new_drawing(
                    ip_type  = zcl_excel_drawing=>type_chart
                    ip_title = 'Stacked CHART BARS WITH 2 SER.' ).
  lo_drawing->graph = lo_bar1_stacked.
  lo_drawing->graph_type = zcl_excel_drawing=>c_graph_bars.

  "Set chart position (anchor 2 cells)
  ls_upper-row = 0.
  ls_upper-col = 1.
  ls_lower-row = 22.
  ls_lower-col = 10.
  lo_drawing->set_position2(
    exporting
      ip_from   = ls_upper
      ip_to     = ls_lower ).

  lo_drawing->set_media(
    exporting
      ip_media_type = zcl_excel_drawing=>c_media_type_xml ).

  lo_worksheet->add_drawing( lo_drawing ).

  " BarChart2 sheet

  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'BarChart2' ).

  " Create global drawing, set type as bar chart, assign chart, set position and media type
  lo_drawing = lo_worksheet->excel->add_new_drawing(
                    ip_type  = zcl_excel_drawing=>type_chart
                    ip_title = 'CHART BARS WITH 1 SERIE' ).
  lo_drawing->graph = lo_bar2.
  lo_drawing->graph_type = zcl_excel_drawing=>c_graph_bars.

  "Set chart position (anchor 2 cells)
  ls_upper-row = 0.
  ls_upper-col = 0.
  ls_lower-row = 30.
  ls_lower-col = 20.
  lo_drawing->set_position2(
    exporting
      ip_from   = ls_upper
      ip_to     = ls_lower ).

  lo_drawing->set_media(
    exporting
      ip_media_type = zcl_excel_drawing=>c_media_type_xml ).

  lo_worksheet->add_drawing( lo_drawing ).

  " LineChart sheet

  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'LineChart' ).

  " Create global drawing, set type as line chart, assign chart, set position and media type
  lo_drawing = lo_worksheet->excel->add_new_drawing(
                    ip_type  = zcl_excel_drawing=>type_chart
                    ip_title = 'CHART LINES' ).
  lo_drawing->graph = lo_line.
  lo_drawing->graph_type = zcl_excel_drawing=>c_graph_line.

  "Set chart position (anchor 2 cells)
  ls_upper-row = 0.
  ls_upper-col = 0.
  ls_lower-row = 30.
  ls_lower-col = 20.
  lo_drawing->set_position2(
    exporting
      ip_from   = ls_upper
      ip_to     = ls_lower ).

  lo_drawing->set_media(
    exporting
      ip_media_type = zcl_excel_drawing=>c_media_type_xml ).

  lo_worksheet->add_drawing( lo_drawing ).

  " Values sheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Values' ).

  " Set values for chart
  lv_value = 1.
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = lv_value ).
  lv_value = 2.
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = lv_value ).
  lv_value = 3.
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = lv_value ).

  " Set labels for chart
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'One' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Two' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'Three' ).

  " Set values for chart
  lv_value = 3.
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = lv_value ).
  lv_value = 2.
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = lv_value ).
  lv_value = -1.
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = lv_value ).

  " Set labels for chart
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 3 ip_value = 'One (Minus)' ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 2 ip_value = 'Two' ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Three' ).

  " Set values for chart
  lv_value = 3.
  lo_worksheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = lv_value ).
  lv_value = 1.
  lo_worksheet->set_cell( ip_column = 'E' ip_row = 2 ip_value = lv_value ).
  lv_value = 2.
  lo_worksheet->set_cell( ip_column = 'E' ip_row = 3 ip_value = lv_value ).

  " Set labels for chart
  lo_worksheet->set_cell( ip_column = 'F' ip_row = 3 ip_value = 'Two' ).
  lo_worksheet->set_cell( ip_column = 'F' ip_row = 2 ip_value = 'One' ).
  lo_worksheet->set_cell( ip_column = 'F' ip_row = 1 ip_value = 'Three' ).


*** Create output
  lcl_output=>output( lo_excel ).
