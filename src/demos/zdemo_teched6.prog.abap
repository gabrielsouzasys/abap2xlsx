*&---------------------------------------------------------------------*
*& Report  ZDEMO_TECHED3
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_teched6.

*******************************
*   Data Object declaration   *
*******************************

data: lo_excel        type ref to zcl_excel,
      lo_excel_writer type ref to zif_excel_writer,
      lo_worksheet    type ref to zcl_excel_worksheet.

data: lo_style_title      type ref to zcl_excel_style,
      lo_drawing          type ref to zcl_excel_drawing,
      lo_range            type ref to zcl_excel_range,
      lo_data_validation  type ref to zcl_excel_data_validation,
      lo_column           type ref to zcl_excel_column,
      lv_style_title_guid type zexcel_cell_style,
      ls_key              type wwwdatatab.

data: lv_file      type xstring,
      lv_bytecount type i,
      lt_file_tab  type solix_tab.

data: lv_full_path      type string,
      lv_workdir        type string,
      lv_file_separator type c.

constants: lv_default_file_name type string value 'TechEd01.xlsx'.

*******************************
* Selection screen management *
*******************************

parameters: p_path type zexcel_export_dir.

at selection-screen on value-request for p_path.
  lv_workdir = p_path.
  cl_gui_frontend_services=>directory_browse( exporting initial_folder  = lv_workdir
                                              changing  selected_folder = lv_workdir ).
  p_path = lv_workdir.

initialization.
  cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.

start-of-selection.

  if p_path is initial.
    p_path = lv_workdir.
  endif.
  cl_gui_frontend_services=>get_file_separator( changing file_separator = lv_file_separator ).
  concatenate p_path lv_file_separator lv_default_file_name into lv_full_path.

*******************************
*    abap2xlsx create XLSX    *
*******************************

  " Create excel instance
  create object lo_excel.

  " Styles
  lo_style_title                   = lo_excel->add_new_style( ).
  lo_style_title->font->bold       = abap_true.
  lo_style_title->font->color-rgb  = zcl_excel_style_color=>c_blue.
  lv_style_title_guid              = lo_style_title->get_guid( ).

  " Get active sheet
  lo_worksheet        = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Demo TechEd' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 5 ip_value = 'TechEd demo' ip_style = lv_style_title_guid ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 7 ip_value = 'Is abap2xlsx simple' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 8 ip_value = 'Is abap2xlsx CooL' ).

  " add logo from SMWO
  lo_drawing = lo_excel->add_new_drawing( ).
  lo_drawing->set_position( ip_from_row = 2
                            ip_from_col = 'B' ).

  ls_key-relid = 'MI'.
  ls_key-objid = 'WBLOGO'.
  lo_drawing->set_media_www( ip_key = ls_key
                             ip_width = 140
                             ip_height = 64 ).

  " assign drawing to the worksheet
  lo_worksheet->add_drawing( lo_drawing ).

  " Add new sheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Values' ).

  " Set values for range
  lo_worksheet->set_cell( ip_row = 4 ip_column = 'A' ip_value = 1 ).
  lo_worksheet->set_cell( ip_row = 5 ip_column = 'A' ip_value = 2 ).
  lo_worksheet->set_cell( ip_row = 6 ip_column = 'A' ip_value = 3 ).
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'A' ip_value = 4 ).
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'A' ip_value = 5 ).

  lo_range            = lo_excel->add_new_range( ).
  lo_range->name      = 'Values'.
  lo_range->set_value( ip_sheet_name    = 'Values'
                       ip_start_column  = 'A'
                       ip_start_row     = 4
                       ip_stop_column   = 'A'
                       ip_stop_row      = 8 ).

  lo_excel->set_active_sheet_index( 1 ).

  " add data validation
  lo_worksheet        = lo_excel->get_active_worksheet( ).

  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
  lo_data_validation->formula1    = 'Values'.
  lo_data_validation->cell_row    = 7.
  lo_data_validation->cell_column = 'C'.
  lo_worksheet->set_cell( ip_row = 7 ip_column = 'C' ip_value = 'Select a value' ).


  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
  lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
  lo_data_validation->formula1    = 'Values'.
  lo_data_validation->cell_row    = 8.
  lo_data_validation->cell_column = 'C'.
  lo_worksheet->set_cell( ip_row = 8 ip_column = 'C' ip_value = 'Select a value' ).

  " add autosize (column width)
  lo_column = lo_worksheet->get_column( ip_column = 'B' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).
  lo_column = lo_worksheet->get_column( ip_column = 'C' ).
  lo_column->set_auto_size( ip_auto_size = abap_true ).

  " Create xlsx stream
  create object lo_excel_writer type zcl_excel_writer_2007.
  lv_file = lo_excel_writer->write_file( lo_excel ).

*******************************
*            Output           *
*******************************

  " Convert to binary
  lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_file ).
  lv_bytecount = xstrlen( lv_file ).

  " Save the file
  cl_gui_frontend_services=>gui_download( exporting bin_filesize = lv_bytecount
                                                    filename     = lv_full_path
                                                    filetype     = 'BIN'
                                           changing data_tab     = lt_file_tab ).
