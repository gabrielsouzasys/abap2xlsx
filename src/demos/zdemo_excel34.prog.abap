*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL2
*& Test Styles for ABAP2XLSX
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel34.

constants: width            type f value '10.14'.
constants: height           type f value '57.75'.

data: current_row type i,
      col         type i,
      col_alpha   type zexcel_cell_column_alpha,
      row         type i,
      row_board   type i,
      colorflag   type i,
      color       type zexcel_style_color_argb,

      lo_column   type ref to zcl_excel_column,
      lo_row      type ref to zcl_excel_row,

      writing1    type string,
      writing2    type string.



data: lo_excel     type ref to zcl_excel,
      lo_worksheet type ref to zcl_excel_worksheet.

constants: gc_save_file_name type string value '34_Static Styles_Chess.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.
  " Creates active sheet
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Spassky_vs_Bronstein' ).

* Header
  current_row = 1.

  add 1 to current_row.
  lo_worksheet->set_cell( ip_row = current_row ip_column = 'B' ip_value = 'White' ).
  lo_worksheet->set_cell( ip_row = current_row ip_column = 'C' ip_value = 'Spassky, Boris V   --  wins in turn 23' ).

  add 1 to current_row.
  lo_worksheet->set_cell( ip_row = current_row ip_column = 'B' ip_value = 'Black' ).
  lo_worksheet->set_cell( ip_row = current_row ip_column = 'C' ip_value = 'Bronstein, David I' ).

  add 1 to current_row.
* Set size of column + Writing above chessboard
  do 8 times.

    writing1 = zcl_excel_common=>convert_column2alpha( sy-index ).
    writing2 =  sy-index .
    row = current_row + sy-index.

    col = sy-index + 1.
    col_alpha = zcl_excel_common=>convert_column2alpha( col ).

* Set size of column
    lo_column = lo_worksheet->get_column( col_alpha ).
    lo_column->set_width( width ).

* Set size of row
    lo_row = lo_worksheet->get_row( row ).
    lo_row->set_row_height( height ).

* Set writing on chessboard
    lo_worksheet->set_cell( ip_row = row
                            ip_column = 'A'
                            ip_value = writing2 ).
    lo_worksheet->change_cell_style(  ip_column               = 'A'
                                      ip_row                  = row
                                      ip_alignment_vertical   = zcl_excel_style_alignment=>c_vertical_center  ).
    lo_worksheet->set_cell( ip_row = row
                            ip_column = 'J'
                            ip_value = writing2 ).
    lo_worksheet->change_cell_style(  ip_column               = 'J'
                                      ip_row                  = row
                                      ip_alignment_vertical   = zcl_excel_style_alignment=>c_vertical_center  ).

    row = current_row + 9.
    lo_worksheet->set_cell( ip_row = current_row
                            ip_column = col_alpha
                            ip_value = writing1 ).
    lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                      ip_row                  = current_row
                                      ip_alignment_horizontal = zcl_excel_style_alignment=>c_horizontal_center ).
    lo_worksheet->set_cell( ip_row = row
                            ip_column = col_alpha
                            ip_value = writing1 ).
    lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                      ip_row                  = row
                                      ip_alignment_horizontal = zcl_excel_style_alignment=>c_horizontal_center ).
  enddo.
  lo_column = lo_worksheet->get_column( 'A' ).
  lo_column->set_auto_size( abap_true ).
  lo_column = lo_worksheet->get_column( 'J' ).
  lo_column->set_auto_size( abap_true ).

* Set win-position
  constants: c_pawn   type string value 'Pawn'.
  constants: c_rook   type string value 'Rook'.
  constants: c_knight type string value 'Knight'.
  constants: c_bishop type string value 'Bishop'.
  constants: c_queen  type string value 'Queen'.
  constants: c_king   type string value 'King'.

  row = current_row + 1.
  lo_worksheet->set_cell( ip_row = row ip_column = 'B' ip_value = c_rook ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'F' ip_value = c_rook ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'G' ip_value = c_knight ).
  row = current_row + 2.
  lo_worksheet->set_cell( ip_row = row ip_column = 'B' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'C' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'D' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'F' ip_value = c_queen ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'H' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'I' ip_value = c_king ).
  row = current_row + 3.
  lo_worksheet->set_cell( ip_row = row ip_column = 'I' ip_value = c_pawn ).
  row = current_row + 4.
  lo_worksheet->set_cell( ip_row = row ip_column = 'D' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'F' ip_value = c_knight ).
  row = current_row + 5.
  lo_worksheet->set_cell( ip_row = row ip_column = 'E' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'F' ip_value = c_queen ).
  row = current_row + 6.
  lo_worksheet->set_cell( ip_row = row ip_column = 'C' ip_value = c_bishop ).
  row = current_row + 7.
  lo_worksheet->set_cell( ip_row = row ip_column = 'B' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'C' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'H' ip_value = c_pawn ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'I' ip_value = c_pawn ).
  row = current_row + 8.
  lo_worksheet->set_cell( ip_row = row ip_column = 'G' ip_value = c_rook ).
  lo_worksheet->set_cell( ip_row = row ip_column = 'H' ip_value = c_king ).

* Set Chessboard
  do 8 times.
    if sy-index <= 3.  " Black
      color = zcl_excel_style_color=>c_black.
    else.
      color = zcl_excel_style_color=>c_white.
    endif.
    row_board = sy-index.
    row = current_row + sy-index.
    do 8 times.
      col = sy-index + 1.
      col_alpha = zcl_excel_common=>convert_column2alpha( col ).
      try.
* Borders around outer limits
          if row_board = 1.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_borders_top_style    = zcl_excel_style_border=>c_border_thick
                                              ip_borders_top_color_rgb =  zcl_excel_style_color=>c_black ).
          endif.
          if row_board = 8.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_borders_down_style    = zcl_excel_style_border=>c_border_thick
                                              ip_borders_down_color_rgb =  zcl_excel_style_color=>c_black ).
          endif.
          if col = 2.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_borders_left_style    = zcl_excel_style_border=>c_border_thick
                                              ip_borders_left_color_rgb =  zcl_excel_style_color=>c_black ).
          endif.
          if col = 9.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_borders_right_style    = zcl_excel_style_border=>c_border_thick
                                              ip_borders_right_color_rgb =  zcl_excel_style_color=>c_black ).
          endif.
* Style for writing
          lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                            ip_row                  = row
                                            ip_font_color_rgb       = color
                                            ip_font_bold            = 'X'
                                            ip_font_size            = 16
                                            ip_alignment_horizontal = zcl_excel_style_alignment=>c_horizontal_center
                                            ip_alignment_vertical   = zcl_excel_style_alignment=>c_vertical_center
                                            ip_fill_filltype        = zcl_excel_style_fill=>c_fill_solid ).
* Color of field
          colorflag = ( row + col ) mod 2.
          if colorflag = 0.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_fill_fgcolor_rgb     = 'FFB5866A'
                                              ip_fill_filltype        = zcl_excel_style_fill=>c_fill_gradient_diagonal135 ).
          else.
            lo_worksheet->change_cell_style(  ip_column               = col_alpha
                                              ip_row                  = row
                                              ip_fill_fgcolor_rgb     = 'FFF5DEBF'
                                              ip_fill_filltype        = zcl_excel_style_fill=>c_fill_gradient_diagonal45 ).
          endif.



        catch zcx_excel .
      endtry.

    enddo.
  enddo.


*** Create output
  lcl_output=>output( lo_excel ).
