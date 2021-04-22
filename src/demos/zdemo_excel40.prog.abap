report zdemo_excel40.


data: lo_excel     type ref to zcl_excel,
      lo_worksheet type ref to zcl_excel_worksheet.

data: lv_row       type zexcel_cell_row,
      lv_col       type i,
      lv_row_char  type char10,
      lv_value     type string,
      ls_fontcolor type zexcel_style_color_argb.

constants: gc_save_file_name type string value '40_Printsettings.xlsx'.
include zdemo_excel_outputopt_incl.



start-of-selection.
  " Creates active sheet
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Demo Printsettings' ).

*--------------------------------------------------------------------*
*  Prepare sheet with trivial data
*    - first 4 columns will have fontocolor set
*    - first 3 rows will  have fontcolor set
*    These marked cells will be used for repeatable rows/columns on printpages
*--------------------------------------------------------------------*
  do 100 times.  " Rows

    lv_row = sy-index .
    write lv_row to lv_row_char.

    do 20 times.

      lv_col = sy-index - 1.
      concatenate sy-abcde+lv_col(1) lv_row_char into lv_value.
      lv_col = sy-index.
      lo_worksheet->set_cell( ip_row    = lv_row
                              ip_column = lv_col
                              ip_value  = lv_value ).

      try.
          if lv_row <= 3.
            lo_worksheet->change_cell_style(  ip_column           = lv_col
                                              ip_row              = lv_row
                                              ip_fill_filltype    = zcl_excel_style_fill=>c_fill_solid
                                              ip_fill_fgcolor_rgb = zcl_excel_style_color=>c_yellow ).
          endif.
          if lv_col <= 4.
            lo_worksheet->change_cell_style(  ip_column           = lv_col
                                              ip_row              = lv_row
                                              ip_font_color_rgb   = zcl_excel_style_color=>c_red ).
          endif.
        catch zcx_excel .
      endtry.

    enddo.



  enddo.


*--------------------------------------------------------------------*
*  Printsettings
*--------------------------------------------------------------------*
  try.
      lo_worksheet->zif_excel_sheet_printsettings~set_print_repeat_columns( iv_columns_from = 'A'
                                                                            iv_columns_to   = 'D' ).
      lo_worksheet->zif_excel_sheet_printsettings~set_print_repeat_rows(    iv_rows_from    = 1
                                                                            iv_rows_to      = 3 ).
    catch zcx_excel .
  endtry.

*** Create output
  lcl_output=>output( lo_excel ).
