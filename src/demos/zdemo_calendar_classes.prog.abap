*&---------------------------------------------------------------------*
*&  Include           ZDEMO_CALENDAR_CLASSES
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*&       Class ZCL_DATE_CALCULATION
*&---------------------------------------------------------------------*
*        Text
*----------------------------------------------------------------------*
class zcl_date_calculation definition.
  public section.
    class-methods: months_between_two_dates
      importing
        i_date_from type datum
        i_date_to   type datum
        i_incl_to   type flag
      exporting
        e_month     type i.
endclass.               "ZCL_DATE_CALCULATION


*----------------------------------------------------------------------*
*       CLASS ZCL_DATE_CALCULATION IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class zcl_date_calculation implementation.
  method months_between_two_dates.
    data: date_to type datum.
    data: begin of datum_von,
            jjjj(4) type n,
            mm(2)   type n,
            tt(2)   type n,
          end of datum_von.

    data: begin of datum_bis,
            jjjj(4) type n,
            mm(2)   type n,
            tt(2)   type n,
          end of datum_bis.

    e_month = 0.

    check not ( i_date_from is initial )
      and not ( i_date_to is initial ).

    date_to = i_date_to.
    if i_incl_to = abap_true.
      date_to = date_to + 1.
    endif.

    datum_von = i_date_from.
    datum_bis = date_to.

    e_month   =   ( datum_bis-jjjj - datum_von-jjjj ) * 12
                + ( datum_bis-mm   - datum_von-mm   ).
  endmethod.                    "MONTHS_BETWEEN_TWO_DATES
endclass.                    "ZCL_DATE_CALCULATION IMPLEMENTATION

*----------------------------------------------------------------------*
*       CLASS zcl_date_calculation_test DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class zcl_date_calculation_test definition for testing
    duration short
    risk level harmless
  .
  private section.
    methods:
      months_between_two_dates for testing.
endclass.                    "zcl_date_calculation_test DEFINITION
*----------------------------------------------------------------------*
*       CLASS zcl_date_calculation_test IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class zcl_date_calculation_test implementation.
  method months_between_two_dates.

    data: date_from type datum value '20120101',
          date_to   type datum value '20121231'.
    data: month type i.

    zcl_date_calculation=>months_between_two_dates(
      exporting
        i_date_from = date_from
        i_date_to   = date_to
        i_incl_to   = abap_true
      importing
        e_month     = month
    ).

    cl_aunit_assert=>assert_equals(
        exp                  = 12    " Data Object with Expected Type
        act                  = month    " Data Object with Current Value
        msg                  = 'Calculated date is wrong'    " Message in Case of Error
    ).

  endmethod.                    "months_between_two_dates
endclass.                    "zcl_date_calculation_test IMPLEMENTATION
*----------------------------------------------------------------------*
*       CLASS zcl_helper DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class zcl_helper definition.
  public section.
    class-methods:
      load_image
        importing
                  filename       type string
        returning value(r_image) type xstring,
      add_calendar
        importing
          i_date_from type datum
          i_date_to   type datum
          i_from_row  type zexcel_cell_row
          i_from_col  type zexcel_cell_column_alpha
          i_day_style type zexcel_cell_style
          i_cw_style  type zexcel_cell_style
        changing
          c_worksheet type ref to zcl_excel_worksheet
        raising
          zcx_excel,
      add_calendar_landscape
        importing
          i_date_from type datum
          i_date_to   type datum
          i_from_row  type zexcel_cell_row
          i_from_col  type zexcel_cell_column_alpha
          i_day_style type zexcel_cell_style
          i_cw_style  type zexcel_cell_style
        changing
          c_worksheet type ref to zcl_excel_worksheet
        raising
          zcx_excel,
      add_a2x_footer
        importing
          i_from_row  type zexcel_cell_row
          i_from_col  type zexcel_cell_column_alpha
        changing
          c_worksheet type ref to zcl_excel_worksheet
        raising
          zcx_excel,
      add_calender_week
        importing
          i_date      type datum
          i_row       type zexcel_cell_row
          i_col       type zexcel_cell_column_alpha
          i_style     type zexcel_cell_style
        changing
          c_worksheet type ref to zcl_excel_worksheet
        raising
          zcx_excel.
endclass.                    "zcl_helper DEFINITION

*----------------------------------------------------------------------*
*       CLASS zcl_helper IMPLEMENTATION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class zcl_helper implementation.
  method load_image.
    "Load samle image
    data: lt_bin type solix_tab,
          lv_len type i.

    call method cl_gui_frontend_services=>gui_upload
      exporting
        filename                = filename
        filetype                = 'BIN'
      importing
        filelength              = lv_len
      changing
        data_tab                = lt_bin
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
        others                  = 19.
    if sy-subrc <> 0.
      message id sy-msgid type sy-msgty number sy-msgno
                 with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    endif.

    call function 'SCMS_BINARY_TO_XSTRING'
      exporting
        input_length = lv_len
      importing
        buffer       = r_image
      tables
        binary_tab   = lt_bin
      exceptions
        failed       = 1
        others       = 2.
    if sy-subrc <> 0.
      message id sy-msgid type sy-msgty number sy-msgno
                 with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    endif.
  endmethod.                    "load_image
  method add_calendar.
    data: day_names     type table of t246.
    data: row          type zexcel_cell_row,
          row_max      type i,
          col_int      type zexcel_cell_column,
          col_max      type i,
          from_col_int type zexcel_cell_column,
          col          type zexcel_cell_column_alpha,
          lo_column    type ref to zcl_excel_column,
          lo_row       type ref to zcl_excel_row.
    data: lv_date type datum,
          value   type string,
          weekday type wotnr,
          weekrow type wotnr value 1,
          day     type i,
          width   type f,
          height  type f.
    data: hyperlink type ref to zcl_excel_hyperlink.

    field-symbols: <day_name> like line of day_names.

    lv_date = i_date_from.
    from_col_int = zcl_excel_common=>convert_column2int( i_from_col ).
    " Add description for Calendar Week
    c_worksheet->set_cell(
      exporting
        ip_column    = i_from_col    " Cell Column
        ip_row       = i_from_row    " Cell Row
        ip_value     = 'CW'(001)    " Cell Value
        ip_style     = i_cw_style
    ).

    " Add Days
    call function 'DAY_NAMES_GET'
      tables
        day_names = day_names.

    loop at day_names assigning <day_name>.
      row     = i_from_row.
      col_int = from_col_int + <day_name>-wotnr.
      col = zcl_excel_common=>convert_column2alpha( col_int ).
      value = <day_name>-langt.
      c_worksheet->set_cell(
        exporting
          ip_column    = col    " Cell Column
          ip_row       = row    " Cell Row
          ip_value     = value    " Cell Value
          ip_style     = i_cw_style
      ).
    endloop.

    while lv_date <= i_date_to.
      day = lv_date+6(2).
      call function 'FIMA_X_DAY_IN_MONTH_COMPUTE'
        exporting
          i_datum        = lv_date
        importing
          e_wochentag_nr = weekday.

      row     = i_from_row + weekrow.
      col_int = from_col_int + weekday.
      col = zcl_excel_common=>convert_column2alpha( col_int ).

      value = day.
      condense value.

      c_worksheet->set_cell(
        exporting
          ip_column    = col    " Cell Column
          ip_row       = row    " Cell Row
          ip_value     = value    " Cell Value
          ip_style     = i_day_style    " Single-Character Indicator
      ).

      if weekday = 7.
        " Add Calender Week
        zcl_helper=>add_calender_week(
          exporting
            i_date      = lv_date
            i_row       = row
            i_col       = i_from_col
            i_style     = i_cw_style
          changing
            c_worksheet = c_worksheet
        ).
        weekrow = weekrow + 1.
      endif.
      lv_date = lv_date + 1.
    endwhile.
    " Add Calender Week
    zcl_helper=>add_calender_week(
      exporting
        i_date      = lv_date
        i_row       = row
        i_col       = i_from_col
        i_style     = i_cw_style
      changing
        c_worksheet = c_worksheet
    ).
    " Add Created with abap2xlsx
    row = row + 2.
    zcl_helper=>add_a2x_footer(
      exporting
        i_from_row  = row
        i_from_col  = i_from_col
      changing
        c_worksheet = c_worksheet
    ).
    col_int = from_col_int.
    col_max = from_col_int + 7.
    while col_int <= col_max.
      col = zcl_excel_common=>convert_column2alpha( col_int ).
      if sy-index = 1.
        width = '5.0'.
      else.
        width = '11.4'.
      endif.
      lo_column = c_worksheet->get_column( col ).
      lo_column->set_width( width ).
      col_int = col_int + 1.
    endwhile.
    row = i_from_row + 1.
    row_max = i_from_row + 6.
    while row <= row_max.
      height = 50.
      lo_row = c_worksheet->get_row( row ).
      lo_row->set_row_height( height ).
      row = row + 1.
    endwhile.
  endmethod.                    "add_calendar
  method add_a2x_footer.
    data: value     type string,
          hyperlink type ref to zcl_excel_hyperlink.

    value = 'Created with abap2xlsx. Find more information at https://github.com/sapmentors/abap2xlsx.'(002).
    hyperlink = zcl_excel_hyperlink=>create_external_link( 'https://github.com/sapmentors/abap2xlsx' ). "#EC NOTEXT
    c_worksheet->set_cell(
      exporting
        ip_column    = i_from_col    " Cell Column
        ip_row       = i_from_row    " Cell Row
        ip_value     = value    " Cell Value
        ip_hyperlink = hyperlink
    ).

  endmethod.                    "add_a2x_footer
  method add_calendar_landscape.
    data: day_names type table of t246.

    data: lv_date type datum,
          day     type i,
          value   type string,
          weekday type wotnr.
    data: row          type zexcel_cell_row,
          from_col_int type zexcel_cell_column,
          col_int      type zexcel_cell_column,
          col          type zexcel_cell_column_alpha.
    data: lo_column type ref to zcl_excel_column,
          lo_row    type ref to zcl_excel_row.

    field-symbols: <day_name> like line of day_names.

    lv_date = i_date_from.
    " Add Days
    call function 'DAY_NAMES_GET'
      tables
        day_names = day_names.

    while lv_date <= i_date_to.
      day = lv_date+6(2).
      call function 'FIMA_X_DAY_IN_MONTH_COMPUTE'
        exporting
          i_datum        = lv_date
        importing
          e_wochentag_nr = weekday.
      " Day name row
      row     = i_from_row.
      col_int = from_col_int + day + 2.
      col = zcl_excel_common=>convert_column2alpha( col_int ).
      read table day_names assigning <day_name>
        with key wotnr = weekday.
      value = <day_name>-kurzt.
      c_worksheet->set_cell(
        exporting
          ip_column    = col    " Cell Column
          ip_row       = row    " Cell Row
          ip_value     = value    " Cell Value
          ip_style     = i_cw_style
      ).

      " Day row
      row     = i_from_row + 1.
      value = day.
      condense value.

      c_worksheet->set_cell(
        exporting
          ip_column    = col    " Cell Column
          ip_row       = row    " Cell Row
          ip_value     = value    " Cell Value
          ip_style     = i_day_style    " Single-Character Indicator
      ).
      " width
      lo_column = c_worksheet->get_column( col ).
      lo_column->set_width( '3.6' ).


      lv_date = lv_date + 1.
    endwhile.
    " Add ABAP2XLSX Footer
    row = i_from_row + 2.
    c_worksheet->set_cell(
      exporting
        ip_column    = col    " Cell Column
        ip_row       = row    " Cell Row
        ip_value     = ' '    " Cell Value
    ).
    lo_row = c_worksheet->get_row( row ).
    lo_row->set_row_height( '5.0' ).
    row = i_from_row + 3.
    zcl_helper=>add_a2x_footer(
      exporting
        i_from_row  = row
        i_from_col  = i_from_col
      changing
        c_worksheet = c_worksheet
    ).

    " Set with for all 31 coulumns
    while day < 32.
      day = day + 1.
      col_int = from_col_int + day + 2.
      col = zcl_excel_common=>convert_column2alpha( col_int ).
      " width
      lo_column = c_worksheet->get_column( col ).
      lo_column->set_width( '3.6' ).
    endwhile.
  endmethod.                    "ADD_CALENDAR_LANDSCAPE

  method add_calender_week.
    data: week     type kweek,
          week_int type i,
          value    type string.
    " Add Calender Week
    call function 'DATE_GET_WEEK'
      exporting
        date = i_date  " Date for which the week should be calculated
      importing
        week = week.    " Week for date (format:YYYYWW)
    value = week+4(2).
    week_int = value.
    value = week_int.
    condense value.
    c_worksheet->set_cell(
      exporting
        ip_column    = i_col    " Cell Column
        ip_row       = i_row    " Cell Row
        ip_value     = value    " Cell Value
        ip_style     = i_style
    ).
  endmethod.                    "add_calender_week
endclass.                    "zcl_helper IMPLEMENTATION
