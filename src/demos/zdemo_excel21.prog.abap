*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL21
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel21.

types:
  begin of t_color_style,
    color type zexcel_style_color_argb,
    style type zexcel_cell_style,
  end of t_color_style.

data: lo_excel        type ref to zcl_excel,
      lo_worksheet    type ref to zcl_excel_worksheet,
      lo_style_filled type ref to zcl_excel_style.

data: color_styles type table of t_color_style.

field-symbols: <color_style> like line of color_styles.

constants: max  type i value 255,
           step type i value 51.

data: red          type i,
      green        type i,
      blue         type i,
      red_hex(1)   type x,
      green_hex(1) type x,
      blue_hex(1)  type x,
      red_str      type string,
      green_str    type string,
      blue_str     type string.

data: color type zexcel_style_color_argb,
      tint  type zexcel_style_color_tint.

data: row     type i,
      row_tmp type i,
      column  type zexcel_cell_column value 1,
      col_str type zexcel_cell_column_alpha.

constants: gc_save_file_name type string value '21_BackgroundColorPicker.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.

  " Creates active sheet
  create object lo_excel.

  while red <= max.
    green = 0.
    while green <= max.
      blue = 0.
      while blue <= max.
        red_hex = red.
        red_str = red_hex.
        green_hex = green.
        green_str = green_hex.
        blue_hex = blue.
        blue_str = blue_hex.
        " Create filled
        concatenate 'FF' red_str green_str blue_str into color.
        append initial line to color_styles assigning <color_style>.
        lo_style_filled                 = lo_excel->add_new_style( ).
        lo_style_filled->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
        lo_style_filled->fill->fgcolor-rgb  = color.
        <color_style>-color = color.
        <color_style>-style = lo_style_filled->get_guid( ).
        blue = blue + step.
      endwhile.
      green = green + step.
    endwhile.
    red = red + step.
  endwhile.
  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Color Picker' ).
  loop at color_styles assigning <color_style>.
    row_tmp = ( max / step + 1 ) * 3.
    if row = row_tmp.
      row = 0.
      column = column + 1.
    endif.
    row = row + 1.
    col_str = zcl_excel_common=>convert_column2alpha( column ).

    " Fill the cell and apply one style
    lo_worksheet->set_cell( ip_column = col_str
                            ip_row    = row
                            ip_value  = <color_style>-color
                            ip_style  = <color_style>-style ).
  endloop.

  row = row + 2.
  tint = '-0.5'.
  do 10 times.
    column = 1.
    do 10 times.
      lo_style_filled                 = lo_excel->add_new_style( ).
      lo_style_filled->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
      lo_style_filled->fill->fgcolor-theme  = sy-index - 1.
      lo_style_filled->fill->fgcolor-tint  = tint.
      <color_style>-style = lo_style_filled->get_guid( ).
      col_str = zcl_excel_common=>convert_column2alpha( column ).
      lo_worksheet->set_cell_style( ip_column = col_str
                                    ip_row    = row
                                    ip_style  = <color_style>-style ).

      add 1 to column.
    enddo.
    add '0.1' to tint.
    add 1 to row.
  enddo.



*** Create output
  lcl_output=>output( lo_excel ).
