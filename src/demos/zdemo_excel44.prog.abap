*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL44
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report  zdemo_excel44.

data: lo_excel_no_line_if_empty     type ref to zcl_excel,
      lo_excel                      type ref to zcl_excel,
      lo_worksheet_no_line_if_empty type ref to zcl_excel_worksheet,
      lo_worksheet                  type ref to zcl_excel_worksheet.

data: lt_field_catalog    type zexcel_t_fieldcatalog.

data: gc_save_file_name type string value '44_iTabEmpty.csv'.
include zdemo_excel_outputopt_incl.

selection-screen begin of block b44 with frame title txt_b44.

* No line if internal table is empty
data: p_mtyfil type flag value abap_true.

selection-screen end of block b44.

initialization.
  txt_b44 = 'Testing empty file option'(b44).

start-of-selection.

  " Creates active sheet
  create object lo_excel_no_line_if_empty.
  create object lo_excel.

  " Get active sheet
  lo_worksheet_no_line_if_empty = lo_excel_no_line_if_empty->get_active_worksheet( ).
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet_no_line_if_empty->set_title( 'Internal table' ).
  lo_worksheet->set_title( 'Internal table' ).

  data lt_test type table of sflight.

  lo_worksheet_no_line_if_empty->bind_table( ip_table            = lt_test
                            iv_no_line_if_empty = p_mtyfil ).

  p_mtyfil = abap_false.
  lo_worksheet->bind_table( ip_table            = lt_test
                            iv_no_line_if_empty = p_mtyfil ).

*** Create output
  lcl_output=>output( exporting cl_excel            = lo_excel_no_line_if_empty
                                iv_writerclass_name = 'ZCL_EXCEL_WRITER_CSV' ).

  gc_save_file_name = '44_iTabNotEmpty.csv'.
  lcl_output=>output( exporting cl_excel            = lo_excel
                              iv_writerclass_name = 'ZCL_EXCEL_WRITER_CSV' ).
