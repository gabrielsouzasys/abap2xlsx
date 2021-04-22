*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL26
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel29.

data: lo_excel        type ref to zcl_excel,
      lo_excel_writer type ref to zif_excel_writer,
      lo_excel_reader type ref to zif_excel_reader.

data: lv_file      type xstring,
      lv_bytecount type i,
      lt_file_tab  type solix_tab.

data: lv_full_path type string,
      lv_filename  type string,
      lv_workdir   type string.

parameters: p_path type zexcel_export_dir obligatory.

at selection-screen on value-request for p_path.

  data: lt_filetable type filetable,
        lv_rc        type i.

  cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

  call method cl_gui_frontend_services=>file_open_dialog
    exporting
      window_title            = 'Select Macro-Enabled Workbook template'
      default_extension       = '*.xlsm'
      file_filter             = 'Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm'
      initial_directory       = lv_workdir
    changing
      file_table              = lt_filetable
      rc                      = lv_rc
    exceptions
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      others                  = 5.
  if sy-subrc <> 0.
    message id sy-msgid type sy-msgty number sy-msgno
               with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  endif.
  read table lt_filetable into lv_filename index 1.
  p_path = lv_filename.

start-of-selection.

  lv_full_path = p_path.

  create object lo_excel_reader type zcl_excel_reader_xlsm.
  create object lo_excel_writer type zcl_excel_writer_xlsm.
  lo_excel = lo_excel_reader->load_file( lv_full_path ).
  lv_file = lo_excel_writer->write_file( lo_excel ).
  replace '.xlsm' in lv_full_path with 'FromReader.xlsm'.

  " Convert to binary
  call function 'SCMS_XSTRING_TO_BINARY'
    exporting
      buffer        = lv_file
    importing
      output_length = lv_bytecount
    tables
      binary_tab    = lt_file_tab.

  " Save the file
  cl_gui_frontend_services=>gui_download( exporting bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           changing data_tab     = lt_file_tab ).
