*&---------------------------------------------------------------------*
*& Report  ZANGRY_BIRDS
*& Just for fun
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zangry_birds.

data: lo_excel        type ref to zcl_excel,
      lo_excel_writer type ref to zif_excel_writer,
      lo_worksheet    type ref to zcl_excel_worksheet,
      lo_border_light type ref to zcl_excel_style_border,
      lo_style_color0 type ref to zcl_excel_style,
      lo_style_color1 type ref to zcl_excel_style,
      lo_style_color2 type ref to zcl_excel_style,
      lo_style_color3 type ref to zcl_excel_style,
      lo_style_color4 type ref to zcl_excel_style,
      lo_style_color5 type ref to zcl_excel_style,
      lo_style_color6 type ref to zcl_excel_style,
      lo_style_color7 type ref to zcl_excel_style,
      lo_style_credit type ref to zcl_excel_style,
      lo_style_link   type ref to zcl_excel_style,
      lo_column       type ref to zcl_excel_column,
      lo_row          type ref to zcl_excel_row,
      lo_hyperlink    type ref to zcl_excel_hyperlink.

data: lv_style_color0_guid type zexcel_cell_style,
      lv_style_color1_guid type zexcel_cell_style,
      lv_style_color2_guid type zexcel_cell_style,
      lv_style_color3_guid type zexcel_cell_style,
      lv_style_color4_guid type zexcel_cell_style,
      lv_style_color5_guid type zexcel_cell_style,
      lv_style_color6_guid type zexcel_cell_style,
      lv_style_color7_guid type zexcel_cell_style,
      lv_style_credit_guid type zexcel_cell_style,
      lv_style_link_guid   type zexcel_cell_style.

data: lv_col_str type zexcel_cell_column_alpha,
      lv_row     type i,
      lv_col     type i,
      lt_mapper  type table of zexcel_cell_style,
      ls_mapper  type zexcel_cell_style.

data: lv_file      type xstring,
      lv_bytecount type i,
      lt_file_tab  type solix_tab.

data: lv_full_path      type string,
      lv_workdir        type string,
      lv_file_separator type c.

constants: lv_default_file_name type string value 'angry_birds.xlsx'.

parameters: p_path type zexcel_export_dir.

at selection-screen on value-request for p_path.
  lv_workdir = p_path.
  cl_gui_frontend_services=>directory_browse( exporting initial_folder  = lv_workdir
                                              changing  selected_folder = lv_workdir ).
  p_path = lv_workdir.

initialization.
  cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

start-of-selection.

  if p_path is initial.
    p_path = lv_workdir.
  endif.
  cl_gui_frontend_services=>get_file_separator( changing file_separator = lv_file_separator ).
  concatenate p_path lv_file_separator lv_default_file_name into lv_full_path.

  " Creates active sheet
  create object lo_excel.

  create object lo_border_light.
  lo_border_light->border_color-rgb = zcl_excel_style_color=>c_white.
  lo_border_light->border_style = zcl_excel_style_border=>c_border_thin.

  " Create color white
  lo_style_color0                      = lo_excel->add_new_style( ).
  lo_style_color0->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color0->fill->fgcolor-rgb   = 'FFFFFFFF'.
  lo_style_color0->borders->allborders = lo_border_light.
  lv_style_color0_guid                 = lo_style_color0->get_guid( ).

  " Create color black
  lo_style_color1                      = lo_excel->add_new_style( ).
  lo_style_color1->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color1->fill->fgcolor-rgb   = 'FF252525'.
  lo_style_color1->borders->allborders = lo_border_light.
  lv_style_color1_guid                 = lo_style_color1->get_guid( ).

  " Create color dark green
  lo_style_color2                      = lo_excel->add_new_style( ).
  lo_style_color2->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color2->fill->fgcolor-rgb   = 'FF75913A'.
  lo_style_color2->borders->allborders = lo_border_light.
  lv_style_color2_guid                 = lo_style_color2->get_guid( ).

  " Create color light green
  lo_style_color3                      = lo_excel->add_new_style( ).
  lo_style_color3->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color3->fill->fgcolor-rgb   = 'FF9DFB73'.
  lo_style_color3->borders->allborders = lo_border_light.
  lv_style_color3_guid                 = lo_style_color3->get_guid( ).

  " Create color green
  lo_style_color4                      = lo_excel->add_new_style( ).
  lo_style_color4->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color4->fill->fgcolor-rgb   = 'FF92CF56'.
  lo_style_color4->borders->allborders = lo_border_light.
  lv_style_color4_guid                 = lo_style_color4->get_guid( ).

  " Create color 2dark green
  lo_style_color5                      = lo_excel->add_new_style( ).
  lo_style_color5->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color5->fill->fgcolor-rgb   = 'FF506228'.
  lo_style_color5->borders->allborders = lo_border_light.
  lv_style_color5_guid                 = lo_style_color5->get_guid( ).

  " Create color yellow
  lo_style_color6                      = lo_excel->add_new_style( ).
  lo_style_color6->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color6->fill->fgcolor-rgb   = 'FFC3E224'.
  lo_style_color6->borders->allborders = lo_border_light.
  lv_style_color6_guid                 = lo_style_color6->get_guid( ).

  " Create color yellow
  lo_style_color7                      = lo_excel->add_new_style( ).
  lo_style_color7->fill->filltype      = zcl_excel_style_fill=>c_fill_solid.
  lo_style_color7->fill->fgcolor-rgb   = 'FFB3C14F'.
  lo_style_color7->borders->allborders = lo_border_light.
  lv_style_color7_guid                 = lo_style_color7->get_guid( ).

  " Credits
  lo_style_credit = lo_excel->add_new_style( ).
  lo_style_credit->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_credit->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
  lo_style_credit->font->size   = 20.
  lv_style_credit_guid = lo_style_credit->get_guid( ).

  " Link
  lo_style_link = lo_excel->add_new_style( ).
  lo_style_link->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_link->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
*  lo_style_link->font->size   = 20.
  lv_style_link_guid = lo_style_link->get_guid( ).

  " Create image map                                                          " line 2
  do 30 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 3
  do 28 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 5 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 4
  do 27 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 5
  do 9 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 15 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 6
  do 7 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 13 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 7
  do 6 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 5 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 11 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 5 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 8
  do 5 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 9
  do 5 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 10
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 11
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 7 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 12
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 13
  do 5 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 8 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 14
  do 5 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 12 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 15
  do 6 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 8 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 16
  do 7 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 7 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 5 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 17
  do 8 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 13 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 18
  do 6 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 23 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 19
  do 5 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 27 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 20
  do 5 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 23 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 21
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 19 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 22
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 17 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 23
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 17 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 8 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 24
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 10 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 25
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 8 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 26
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color6_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 27
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color6_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 28
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color6_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 29
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 5 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 30
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 31
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color4_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 32
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 8 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color5_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 33
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 34
  do 3 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 10 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 35
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 11 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 36
  do 4 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 10 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 11 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 37
  do 5 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 10 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color7_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 11 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 38
  do 6 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 10 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 11 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 39
  do 7 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 22 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 1 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 40
  do 7 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 17 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 41
  do 8 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 3 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 15 times. append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 42
  do 9 times.  append lv_style_color0_guid to lt_mapper. enddo.
  do 5 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 9 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 2 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 43
  do 11 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 5 times.  append lv_style_color3_guid to lt_mapper. enddo.
  do 7 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 44
  do 13 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 6 times.  append lv_style_color1_guid to lt_mapper. enddo.
  do 4 times.  append lv_style_color2_guid to lt_mapper. enddo.
  do 8 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 45
  do 16 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 13 times. append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape
  " line 46
  do 18 times. append lv_style_color0_guid to lt_mapper. enddo.
  do 8 times.  append lv_style_color1_guid to lt_mapper. enddo.
  append initial line to lt_mapper. " escape

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Angry Birds' ).

  lv_row = 1.
  lv_col = 1.

  loop at lt_mapper into ls_mapper.
    lv_col_str = zcl_excel_common=>convert_column2alpha( lv_col ).
    if ls_mapper is initial.
      lo_row = lo_worksheet->get_row( ip_row = lv_row ).
      lo_row->set_row_height( ip_row_height = 8 ).
      lv_col = 1.
      lv_row = lv_row + 1.
      continue.
    endif.
    lo_worksheet->set_cell( ip_column = lv_col_str
                            ip_row    = lv_row
                            ip_value  = space
                            ip_style  = ls_mapper ).
    lv_col = lv_col + 1.

    lo_column = lo_worksheet->get_column( ip_column = lv_col_str ).
    lo_column->set_width( ip_width = 2 ).
  endloop.

  lo_worksheet->set_show_gridlines( i_show_gridlines = abap_false ).

  lo_worksheet->set_cell( ip_column = 'AP'
                          ip_row    = 15
                          ip_value  = 'Created with abap2xlsx'
                          ip_style  = lv_style_credit_guid ).

  lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = 'https://sapmentors.github.io/abap2xlsx' ).
  lo_worksheet->set_cell( ip_column    = 'AP'
                          ip_row       = 24
                          ip_value     = 'https://sapmentors.github.io/abap2xlsx'
                          ip_style     = lv_style_link_guid
                          ip_hyperlink = lo_hyperlink ).

  lo_column = lo_worksheet->get_column( ip_column = 'AP' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_worksheet->set_merge( ip_row = 15 ip_column_start = 'AP' ip_row_to = 22 ip_column_end = 'AR' ).
  lo_worksheet->set_merge( ip_row = 24 ip_column_start = 'AP' ip_row_to = 26 ip_column_end = 'AR' ).

  create object lo_excel_writer type zcl_excel_writer_2007.
  lv_file = lo_excel_writer->write_file( lo_excel ).

  " Convert to binary
  call function 'SCMS_XSTRING_TO_BINARY'
    exporting
      buffer        = lv_file
    importing
      output_length = lv_bytecount
    tables
      binary_tab    = lt_file_tab.
*  " This method is only available on AS ABAP > 6.40
*  lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
*  lv_bytecount = xstrlen( lv_file ).

  " Save the file
  cl_gui_frontend_services=>gui_download( exporting bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           changing data_tab     = lt_file_tab ).
