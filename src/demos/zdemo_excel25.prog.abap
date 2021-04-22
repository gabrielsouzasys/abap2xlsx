*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL25
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel25.

data: lo_excel        type ref to zcl_excel,
      lo_excel_writer type ref to zif_excel_writer,
      lo_worksheet    type ref to zcl_excel_worksheet,
      lo_exception    type ref to cx_root.

data: lv_file                 type xstring.

constants: lv_file_name type string value '25_HelloWorld.xlsx'.
data: lv_default_file_name type string.
data: lv_error type string.

call function 'FILE_GET_NAME_USING_PATH'
  exporting
    logical_path        = 'LOCAL_TEMPORARY_FILES'  " Logical path'
    file_name           = lv_file_name    " File name
  importing
    file_name_with_path = lv_default_file_name.    " File name with path
" Creates active sheet
create object lo_excel.

" Get active sheet
lo_worksheet = lo_excel->get_active_worksheet( ).
lo_worksheet->set_title( ip_title = 'Sheet1' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Hello world' ).

create object lo_excel_writer type zcl_excel_writer_2007.
lv_file = lo_excel_writer->write_file( lo_excel ).

try.
    open dataset lv_default_file_name for output in binary mode.
    transfer lv_file  to lv_default_file_name.
    close dataset lv_default_file_name.
  catch cx_root into lo_exception.
    lv_error = lo_exception->get_text( ).
    message lv_error type 'I'.
endtry.
