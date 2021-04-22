*&---------------------------------------------------------------------*
*& Report  ZDEMO_CALENDAR
*& abap2xlsx Demo: Create Calendar with Pictures
*&---------------------------------------------------------------------*
*& This report creates a monthly calendar in the specified date range.
*& Each month is put on a seperate worksheet. The pictures for each
*& month can be specified in a tab delimited UTF-8 encoded file called
*& "Calendar.txt" which is saved in the Export Directory.
*& By default this is the SAP Workdir. The file contains 3 fields:
*&
*& Month (with leading 0)
*& Image Filename
*& Image Description
*& URL for the Description
*&
*& The Images should be landscape JPEG's with a 3:2 ratio and min.
*& 450 pixel height. They must also be saved in the Export Directory.
*& In my tests I've discovered a limit of 20 MB in the
*& cl_gui_frontend_services=>gui_download method. So keep your images
*& smaller or change to a server export using OPEN DATASET.
*&---------------------------------------------------------------------*

report zdemo_calendar.

type-pools: abap.
constants: gc_save_file_name type string value 'Calendar.xlsx'.
include zdemo_excel_outputopt_incl.
include zdemo_calendar_classes.

data: lv_workdir        type string.

parameters: p_from type dfrom,
            p_to   type dto.

selection-screen begin of block orientation with frame title orient.
parameters: p_portr type flag radiobutton group orie,
            p_lands type flag radiobutton group orie default 'X'.
selection-screen end of block orientation.

initialization.
  cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = lv_workdir ).
  cl_gui_cfw=>flush( ).
  p_path = lv_workdir.
  orient = 'Orientation'(000).
  p_from = |{ sy-datum(4) }0101|.
  p_to   = |{ sy-datum(4) }1231|.

start-of-selection.

  data: lo_excel     type ref to zcl_excel,
        lo_worksheet type ref to zcl_excel_worksheet,
        lo_column    type ref to zcl_excel_column,
        lo_row       type ref to zcl_excel_row,
        lo_hyperlink type ref to zcl_excel_hyperlink,
        lo_drawing   type ref to zcl_excel_drawing.

  data: lo_style_month      type ref to zcl_excel_style,
        lv_style_month_guid type zexcel_cell_style.
  data: lo_style_border      type ref to zcl_excel_style,
        lo_border_dark       type ref to zcl_excel_style_border,
        lv_style_border_guid type zexcel_cell_style.
  data: lo_style_center      type ref to zcl_excel_style,
        lv_style_center_guid type zexcel_cell_style.

  data: lv_full_path      type string,
        image_descr_path  type string,
        lv_file_separator type c.
  data: lv_content  type xstring,
        width       type i,
        lv_height   type i,
        lv_from_row type zexcel_cell_row.

  data: month      type i,
        month_nr   type fcmnr,
        count      type i value 1,
        title      type zexcel_sheet_title,
        value      type string,
        image_path type string,
        date_from  type datum,
        date_to    type datum,
        row        type zexcel_cell_row,
        to_row     type zexcel_cell_row,
        to_col     type zexcel_cell_column_alpha,
        to_col_end type zexcel_cell_column_alpha,
        to_col_int type i.

  data: month_names type table of t247.
  field-symbols: <month_name> like line of month_names.

  types: begin of tt_datatab,
           month_nr type fcmnr,
           filename type string,
           descr    type string,
           url      type string,
         end of tt_datatab.

  data: image_descriptions type table of tt_datatab.
  field-symbols: <img_descr> like line of image_descriptions.

  constants: lv_default_file_name type string       value 'Calendar', "#EC NOTEXT
             c_from_row_portrait  type zexcel_cell_row value 28,
             c_from_row_landscape type zexcel_cell_row value 38,
             from_col             type zexcel_cell_column_alpha value 'C',
             c_height_portrait    type i              value 450,   " Image Height in Portrait Mode
             c_height_landscape   type i              value 670,   " Image Height in Landscape Mode
             c_factor             type f                        value '1.5'. " Image Ratio, default 3:2

  if p_path is initial.
    p_path = lv_workdir.
  endif.
  cl_gui_frontend_services=>get_file_separator( changing file_separator = lv_file_separator ).
  concatenate p_path lv_file_separator lv_default_file_name '.xlsx' into lv_full_path. "#EC NOTEXT

  " Read Image Names for Month and Description
  concatenate p_path lv_file_separator lv_default_file_name '.txt' into image_descr_path. "#EC NOTEXT
  cl_gui_frontend_services=>gui_upload(
    exporting
      filename                = image_descr_path   " Name of file
      filetype                = 'ASC'    " File Type (ASCII, Binary)
      has_field_separator     = 'X'
      read_by_line            = 'X'    " File Written Line-By-Line to the Internal Table
      codepage                = '4110'
    changing
      data_tab                = image_descriptions    " Transfer table for file contents
    exceptions
      file_open_error         = 1
      file_read_error         = 2
      no_batch                = 3
      gui_refuse_filetransfer = 4
      invalid_type            = 5
      no_authority            = 6
      unknown_error           = 7
      bad_data_format         = 8
      header_not_allowed      = 9
      separator_not_allowed   = 10
      header_too_long         = 11
      unknown_dp_error        = 12
      access_denied           = 13
      dp_out_of_memory        = 14
      disk_full               = 15
      dp_timeout              = 16
      not_supported_by_gui    = 17
      error_no_gui            = 18
      others                  = 19
  ).
  if sy-subrc <> 0.
    message id sy-msgid type sy-msgty number sy-msgno
               with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  endif.

  " Creates active sheet
  create object lo_excel.

  " Create Styles
  " Create an underline double style
  lo_style_month                        = lo_excel->add_new_style( ).
  " lo_style_month->font->underline       = abap_true.
  " lo_style_month->font->underline_mode  = zcl_excel_style_font=>c_underline_single.
  lo_style_month->font->name            = zcl_excel_style_font=>c_name_roman.
  lo_style_month->font->scheme          = zcl_excel_style_font=>c_scheme_none.
  lo_style_month->font->family          = zcl_excel_style_font=>c_family_roman.
  lo_style_month->font->bold            = abap_true.
  lo_style_month->font->size            = 36.
  lv_style_month_guid                   = lo_style_month->get_guid( ).
  " Create border object
  create object lo_border_dark.
  lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
  lo_border_dark->border_style     = zcl_excel_style_border=>c_border_thin.
  "Create style with border
  lo_style_border                         = lo_excel->add_new_style( ).
  lo_style_border->borders->allborders    = lo_border_dark.
  lo_style_border->alignment->horizontal  = zcl_excel_style_alignment=>c_horizontal_right.
  lo_style_border->alignment->vertical    = zcl_excel_style_alignment=>c_vertical_top.
  lv_style_border_guid                    = lo_style_border->get_guid( ).
  "Create style alignment center
  lo_style_center                         = lo_excel->add_new_style( ).
  lo_style_center->alignment->horizontal  = zcl_excel_style_alignment=>c_horizontal_center.
  lo_style_center->alignment->vertical    = zcl_excel_style_alignment=>c_vertical_top.
  lv_style_center_guid                    = lo_style_center->get_guid( ).

  " Get Month Names
  call function 'MONTH_NAMES_GET'
    tables
      month_names = month_names.

  zcl_date_calculation=>months_between_two_dates(
    exporting
      i_date_from = p_from
      i_date_to   = p_to
      i_incl_to   = abap_true
    importing
      e_month     = month
  ).

  date_from = p_from.

  while count <= month.
    if count = 1.
      " Get active sheet
      lo_worksheet = lo_excel->get_active_worksheet( ).
    else.
      lo_worksheet = lo_excel->add_new_worksheet( ).
    endif.

    lo_worksheet->zif_excel_sheet_properties~selected = zif_excel_sheet_properties=>c_selected.

    title = count.
    value = count.
    condense title.
    condense value.
    lo_worksheet->set_title( title ).
    lo_worksheet->set_print_gridlines( abap_false ).
    lo_worksheet->sheet_setup->paper_size = zcl_excel_sheet_setup=>c_papersize_a4.
    lo_worksheet->sheet_setup->horizontal_centered = abap_true.
    lo_worksheet->sheet_setup->vertical_centered   = abap_true.
    lo_column = lo_worksheet->get_column( 'A' ).
    lo_column->set_width( '1.0' ).
    lo_column = lo_worksheet->get_column( 'B' ).
    lo_column->set_width( '2.0' ).
    if p_lands = abap_true.
      lo_worksheet->sheet_setup->orientation = zcl_excel_sheet_setup=>c_orientation_landscape.
      lv_height = c_height_landscape.
      lv_from_row = c_from_row_landscape.
      lo_worksheet->sheet_setup->margin_top    = '0.10'.
      lo_worksheet->sheet_setup->margin_left   = '0.10'.
      lo_worksheet->sheet_setup->margin_right  = '0.10'.
      lo_worksheet->sheet_setup->margin_bottom = '0.10'.
    else.
      lo_column = lo_worksheet->get_column( 'K' ).
      lo_column->set_width( '3.0' ).
      lo_worksheet->sheet_setup->margin_top    = '0.80'.
      lo_worksheet->sheet_setup->margin_left   = '0.55'.
      lo_worksheet->sheet_setup->margin_right  = '0.05'.
      lo_worksheet->sheet_setup->margin_bottom = '0.30'.
      lv_height = c_height_portrait.
      lv_from_row = c_from_row_portrait.
    endif.

    " Add Month Name
    month_nr = date_from+4(2).
    if p_portr = abap_true.
      read table month_names with key mnr = month_nr assigning <month_name>.
      concatenate <month_name>-ltx ` ` date_from(4) into value.
      row = lv_from_row - 2.
      to_col = from_col.
    else.
      row = lv_from_row - 1.
      to_col_int = zcl_excel_common=>convert_column2int( from_col ) + 32.
      to_col = zcl_excel_common=>convert_column2alpha( to_col_int ).
      to_col_int = to_col_int + 1.
      to_col_end = zcl_excel_common=>convert_column2alpha( to_col_int ).
      concatenate month_nr '/' date_from+2(2) into value.
      to_row = row + 2.
      lo_worksheet->set_merge(
        exporting
          ip_column_start = to_col  " Cell Column Start
          ip_column_end   = to_col_end    " Cell Column End
          ip_row          = row       " Cell Row
          ip_row_to       = to_row    " Cell Row
      ).
    endif.
    lo_worksheet->set_cell(
      exporting
        ip_column    = to_col " Cell Column
        ip_row       = row      " Cell Row
        ip_value     = value    " Cell Value
        ip_style     = lv_style_month_guid
    ).

*    to_col_int = zcl_excel_common=>convert_column2int( from_col ) + 7.
*    to_col = zcl_excel_common=>convert_column2alpha( to_col_int ).
*
*    lo_worksheet->set_merge(
*      EXPORTING
*        ip_column_start = from_col  " Cell Column Start
*        ip_column_end   = to_col    " Cell Column End
*        ip_row          = row       " Cell Row
*        ip_row_to       = row       " Cell Row
*    ).

    " Add drawing from a XSTRING read from a file
    unassign <img_descr>.
    read table image_descriptions with key month_nr = month_nr assigning <img_descr>.
    if <img_descr> is assigned.
      value = <img_descr>-descr.
      if p_portr = abap_true.
        row = lv_from_row - 3.
      else.
        row = lv_from_row - 2.
      endif.
      if not <img_descr>-url is initial.
        lo_hyperlink = zcl_excel_hyperlink=>create_external_link( <img_descr>-url ).
        lo_worksheet->set_cell(
          exporting
            ip_column    = from_col " Cell Column
            ip_row       = row      " Cell Row
            ip_value     = value    " Cell Value
            ip_hyperlink = lo_hyperlink
        ).
      else.
        lo_worksheet->set_cell(
          exporting
            ip_column    = from_col " Cell Column
            ip_row       = row      " Cell Row
            ip_value     = value    " Cell Value
        ).
      endif.
      lo_row = lo_worksheet->get_row( row ).
      lo_row->set_row_height( '22.0' ).

      " In Landscape mode the row between the description and the
      " dates should be not so high
      if p_lands = abap_true.
        row = lv_from_row - 3.
        lo_worksheet->set_cell(
          exporting
            ip_column    = from_col " Cell Column
            ip_row       = row      " Cell Row
            ip_value     = ' '      " Cell Value
        ).
        lo_row = lo_worksheet->get_row( row ).
        lo_row->set_row_height( '7.0' ).
        row = lv_from_row - 1.
        lo_row = lo_worksheet->get_row( row ).
        lo_row->set_row_height( '5.0' ).
      endif.

      concatenate p_path lv_file_separator <img_descr>-filename into image_path.
      lo_drawing = lo_excel->add_new_drawing( ).
      lo_drawing->set_position( ip_from_row = 1
                                ip_from_col = 'B' ).

      lv_content = zcl_helper=>load_image( image_path ).
      width = lv_height * c_factor.
      lo_drawing->set_media( ip_media      = lv_content
                             ip_media_type = zcl_excel_drawing=>c_media_type_jpg
                             ip_width      = width
                             ip_height     = lv_height  ).
      lo_worksheet->add_drawing( lo_drawing ).
    endif.

    " Add Calendar
*    CALL FUNCTION 'SLS_MISC_GET_LAST_DAY_OF_MONTH'
*      EXPORTING
*        day_in            = date_from
*      IMPORTING
*        last_day_of_month = date_to.
    date_to = date_from.
    date_to+6(2) = '01'.              " First of month
    add 31 to date_to.                " Somewhere in following month
    date_to = date_to - date_to+6(2). " Last of month
    if p_portr = abap_true.
      zcl_helper=>add_calendar(
        exporting
          i_date_from = date_from
          i_date_to   = date_to
          i_from_row  = lv_from_row
          i_from_col  = from_col
          i_day_style = lv_style_border_guid
          i_cw_style  = lv_style_center_guid
        changing
          c_worksheet = lo_worksheet
      ).
    else.
      zcl_helper=>add_calendar_landscape(
        exporting
          i_date_from = date_from
          i_date_to   = date_to
          i_from_row  = lv_from_row
          i_from_col  = from_col
          i_day_style = lv_style_border_guid
          i_cw_style  = lv_style_center_guid
        changing
          c_worksheet = lo_worksheet
      ).
    endif.
    count = count + 1.
    date_from = date_to + 1.
  endwhile.

  lo_excel->set_active_sheet_index_by_name( '1' ).
*** Create output
  lcl_output=>output( lo_excel ).
