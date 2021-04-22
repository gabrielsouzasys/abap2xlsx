*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL15
*&
*&---------------------------------------------------------------------*
*& 2010-10-30, Gregor Wolf:
*& Added the functionality to ouput the read table content
*& 2011-12-19, Shahrin Shahrulzaman:
*& Added the functionality to have multiple input and output files
*&---------------------------------------------------------------------*

report zdemo_excel15.

type-pools: abap.

types:
  begin of t_demo_excel15,
    input type string,
  end of t_demo_excel15.

constants: sheet_with_date_formats type string value '24_Sheets_with_different_default_date_formats.xlsx'.

data: excel           type ref to zcl_excel,
      lo_excel_writer type ref to zif_excel_writer,
      reader          type ref to zif_excel_reader.

data: ex  type ref to zcx_excel,
      msg type string.

data: lv_file      type xstring,
      lv_bytecount type i,
      lt_file_tab  type solix_tab.

data: lv_workdir        type string,
      output_file_path  type string,
      input_file_path   type string,
      lv_file_separator type c.

data: worksheet      type ref to zcl_excel_worksheet,
      highest_column type zexcel_cell_column,
      highest_row    type int4,
      column         type zexcel_cell_column value 1,
      col_str        type zexcel_cell_column_alpha,
      row            type int4               value 1,
      value          type zexcel_cell_value,
      converted_date type d.

data:
      lt_files       type table of t_demo_excel15.
field-symbols: <wa_files> type t_demo_excel15.

parameters: p_path  type zexcel_export_dir,
            p_noout type abap_bool default abap_true.


at selection-screen on value-request for p_path.
  lv_workdir = p_path.
  cl_gui_frontend_services=>directory_browse( exporting initial_folder  = lv_workdir
                                              changing  selected_folder = lv_workdir ).
  p_path = lv_workdir.

initialization.
  cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input  = '01_HelloWorld.xlsx'.
  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input  = '02_Styles.xlsx'.
  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input  = '03_iTab.xlsx'.
  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input  = '04_Sheets.xlsx'.
  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input = '05_Conditional.xlsx'.
  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input = '07_ConditionalAll.xlsx'.
  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input  = '08_Range.xlsx'.
  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input  = '13_MergedCells.xlsx'.
  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input  = sheet_with_date_formats.
  append initial line to lt_files assigning <wa_files>.
  <wa_files>-input  = '31_AutosizeWithDifferentFontSizes.xlsx'.

start-of-selection.

  if p_path is initial.
    p_path = lv_workdir.
  endif.
  cl_gui_frontend_services=>get_file_separator( changing file_separator = lv_file_separator ).

  loop at lt_files assigning <wa_files>.
    concatenate p_path lv_file_separator <wa_files>-input into input_file_path.
    concatenate p_path lv_file_separator '15_' <wa_files>-input into output_file_path.
    replace '.xlsx' in output_file_path with 'FromReader.xlsx'.

    try.
        create object reader type zcl_excel_reader_2007.
        excel = reader->load_file( input_file_path ).

        if p_noout eq abap_false.
          worksheet = excel->get_active_worksheet( ).
          highest_column = worksheet->get_highest_column( ).
          highest_row    = worksheet->get_highest_row( ).

          write: / 'Filename ', <wa_files>-input.
          write: / 'Highest column: ', highest_column, 'Highest row: ', highest_row.
          write: /.

          while row <= highest_row.
            while column <= highest_column.
              col_str = zcl_excel_common=>convert_column2alpha( column ).
              worksheet->get_cell(
                exporting
                  ip_column = col_str
                  ip_row    = row
                importing
                  ep_value = value
              ).
              write: value.
              column = column + 1.
            endwhile.
            write: /.
            column = 1.
            row = row + 1.
          endwhile.
          if <wa_files>-input = sheet_with_date_formats.
            worksheet->get_cell(
              exporting
                ip_column = 'A'
                ip_row = 4
              importing
                ep_value = value
            ).
            write: / 'Date value using get_cell: ', value.
            converted_date = zcl_excel_common=>excel_string_to_date( ip_value = value ).
            write: / 'Converted date: ', converted_date.
          endif.
        endif.
        create object lo_excel_writer type zcl_excel_writer_2007.
        lv_file = lo_excel_writer->write_file( excel ).

        " Convert to binary
        call function 'SCMS_XSTRING_TO_BINARY'
          exporting
            buffer        = lv_file
          importing
            output_length = lv_bytecount
          tables
            binary_tab    = lt_file_tab.
*    " This method is only available on AS ABAP > 6.40
*    lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
*    lv_bytecount = xstrlen( lv_file ).

        " Save the file
        cl_gui_frontend_services=>gui_download( exporting bin_filesize = lv_bytecount
                                                          filename     = output_file_path
                                                          filetype     = 'BIN'
                                                 changing data_tab     = lt_file_tab ).


      catch zcx_excel into ex.    " Exceptions for ABAP2XLSX
        msg = ex->get_text( ).
        write: / msg.
    endtry.
  endloop.
