*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL1
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel31.

data: lo_excel     type ref to zcl_excel,
      lo_worksheet type ref to zcl_excel_worksheet,
      lo_hyperlink type ref to zcl_excel_hyperlink,
      lo_column    type ref to zcl_excel_column.


data: fieldval            type text80,
      row                 type i,
      style_column_a      type ref to zcl_excel_style,
      style_column_a_guid type zexcel_cell_style,
      style_column_b      type ref to zcl_excel_style,
      style_column_b_guid type zexcel_cell_style,
      style_column_c      type ref to zcl_excel_style,
      style_column_c_guid type zexcel_cell_style,
      style_font          type ref to zcl_excel_style_font.

constants: gc_save_file_name type string value '31_AutosizeWithDifferentFontSizes.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.

  create object lo_excel.
  " Use active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Regular Font' ).

  style_column_a              = lo_excel->add_new_style( ).
  style_column_a->font->size  = 32.                            " quite large
  style_column_a_guid         = style_column_a->get_guid( ).

  style_column_c              = lo_excel->add_new_style( ).
  style_column_c->font->size  = 16.                            " not so large
  style_column_c_guid         = style_column_c->get_guid( ).


  do 20 times.
    row = sy-index.
    clear fieldval.
    do sy-index times.
      concatenate fieldval 'X' into fieldval.
    enddo.
    lo_worksheet->set_cell( ip_column = 'A' ip_row = row ip_value = fieldval ip_style = style_column_a_guid ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = row ip_value = fieldval ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = row ip_value = fieldval ip_style = style_column_c_guid ).
  enddo.

  lo_column = lo_worksheet->get_column( 'A' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_column = lo_worksheet->get_column( 'B' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_column = lo_worksheet->get_column( 'C' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).

  " Add sheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Bold Font' ).

  style_column_a              = lo_excel->add_new_style( ).
  style_column_a->font->size  = 32.                            " quite large
  style_column_a->font->bold  = abap_true.
  style_column_a_guid         = style_column_a->get_guid( ).

  style_column_b              = lo_excel->add_new_style( ).
  style_column_b->font->bold  = abap_true.
  style_column_b_guid         = style_column_b->get_guid( ).

  style_column_c              = lo_excel->add_new_style( ).
  style_column_c->font->size  = 16.                            " not so large
  style_column_c->font->bold  = abap_true.
  style_column_c_guid         = style_column_c->get_guid( ).

  do 20 times.
    row = sy-index.
    clear fieldval.
    do sy-index times.
      concatenate fieldval 'X' into fieldval.
    enddo.
    lo_worksheet->set_cell( ip_column = 'A' ip_row = row ip_value = fieldval ip_style = style_column_a_guid ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = row ip_value = fieldval ip_style = style_column_b_guid ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = row ip_value = fieldval ip_style = style_column_c_guid ).
  enddo.

  lo_column = lo_worksheet->get_column( 'A' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_column = lo_worksheet->get_column( 'B' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_column = lo_worksheet->get_column( 'C' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).

  " Add sheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Italic Font' ).

  style_column_a               = lo_excel->add_new_style( ).
  style_column_a->font->size   = 32.                            " quite large
  style_column_a->font->italic = abap_true.
  style_column_a_guid          = style_column_a->get_guid( ).

  style_column_b               = lo_excel->add_new_style( ).
  style_column_b->font->italic = abap_true.
  style_column_b_guid          = style_column_b->get_guid( ).

  style_column_c               = lo_excel->add_new_style( ).
  style_column_c->font->size   = 16.                            " not so large
  style_column_c->font->italic = abap_true.
  style_column_c_guid          = style_column_c->get_guid( ).

  do 20 times.
    row = sy-index.
    clear fieldval.
    do sy-index times.
      concatenate fieldval 'X' into fieldval.
    enddo.
    lo_worksheet->set_cell( ip_column = 'A' ip_row = row ip_value = fieldval ip_style = style_column_a_guid ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = row ip_value = fieldval ip_style = style_column_b_guid ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = row ip_value = fieldval ip_style = style_column_c_guid ).
  enddo.

  lo_column = lo_worksheet->get_column( 'A' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_column = lo_worksheet->get_column( 'B' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_column = lo_worksheet->get_column( 'C' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).

  " Add sheet for merged cells
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Merged cells' ).

  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'This is a very long header text' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'Some data' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'Some more data' ).

  lo_worksheet->set_merge(
    exporting
      ip_column_start = 'A'
      ip_column_end   = 'C'
      ip_row          = 1 ).

  lo_column = lo_worksheet->get_column( 'A' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).

  lo_excel->set_active_sheet_index( i_active_worksheet = 1 ).

*** Create output
  lcl_output=>output( lo_excel ).
