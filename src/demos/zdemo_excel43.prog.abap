*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL43
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel43.

"
"Locally created Structure, which should be equal to the excels structure
"
types: begin of lty_excel_s,
         dummy type dummy.
types: end of lty_excel_s.

data lt_tab type table of lty_excel_s.
data: lt_filetable type filetable,
      ls_filetable type file_table.
data lv_subrc type i.
data: lo_excel     type ref to zcl_excel,
      lo_reader    type ref to zif_excel_reader,
      lo_worksheet type ref to zcl_excel_worksheet,
      lo_salv      type ref to cl_salv_table.

"
"Ask User to choose a path
"
cl_gui_frontend_services=>file_open_dialog( exporting window_title = 'Excel selection'
                                                      file_filter = '*.xlsx'
                                                      multiselection = abap_false
                                            changing  file_table = lt_filetable " Tabelle, die selektierte Dateien enthält
                                                      rc = lv_subrc
                                            exceptions file_open_dialog_failed = 1
                                                       cntl_error = 2
                                                       error_no_gui = 3
                                                       not_supported_by_gui = 4
                                                       others = 5 ).
if sy-subrc <> 0.
  message id sy-msgid type sy-msgty number sy-msgno
  with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
else.
  create object lo_reader type zcl_excel_reader_2007.
  loop at lt_filetable into ls_filetable.
    lo_excel =  lo_reader->load_file( ls_filetable-filename ).
    lo_worksheet = lo_excel->get_worksheet_by_index( iv_index = 1 ).
    lo_worksheet->get_table( importing et_table = lt_tab ).
  endloop.
endif.
"
"Do the presentation stuff
"

cl_salv_table=>factory( importing r_salv_table = lo_salv
                        changing t_table = lt_tab ).
lo_salv->display( ).
