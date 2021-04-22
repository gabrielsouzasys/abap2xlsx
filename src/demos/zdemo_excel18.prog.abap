*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL18
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel18.

data: lo_excel                 type ref to zcl_excel,
      lo_worksheet             type ref to zcl_excel_worksheet,
      lv_style_protection_guid type zexcel_cell_style.


constants: gc_save_file_name type string value '18_BookProtection.xlsx'.
include zdemo_excel_outputopt_incl.


start-of-selection.

  create object lo_excel.

  " Get active sheet
  lo_excel->zif_excel_book_protection~protected     = zif_excel_book_protection=>c_protected.
  lo_excel->zif_excel_book_protection~lockrevision  = zif_excel_book_protection=>c_locked.
  lo_excel->zif_excel_book_protection~lockstructure = zif_excel_book_protection=>c_locked.
  lo_excel->zif_excel_book_protection~lockwindows   = zif_excel_book_protection=>c_locked.

  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'C' ip_value = 'This cell is unlocked' ip_style = lv_style_protection_guid ).

*** Create output
  lcl_output=>output( lo_excel ).
