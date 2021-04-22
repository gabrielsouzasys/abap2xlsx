*&---------------------------------------------------------------------*
*& Report  ZABAP2XLSX_DEMO_SHOW
*&---------------------------------------------------------------------*
report zabap2xlsx_demo_show.


*----------------------------------------------------------------------*
*       CLASS lcl_perform DEFINITION
*----------------------------------------------------------------------*
class lcl_perform definition create private.
  public section.
    class-methods: setup_objects,
      collect_reports,

      handle_nav for event double_click of cl_gui_alv_grid
        importing e_row.

  private section.
    types: begin of ty_reports,
             progname    type reposrc-progname,
             sort        type reposrc-progname,
             description type repti,
             filename    type string,
           end of ty_reports.

    class-data:
      lo_grid     type ref to cl_gui_alv_grid,
      lo_text     type ref to cl_gui_textedit,
      cl_document type ref to i_oi_document_proxy,

      t_reports   type standard table of ty_reports with non-unique default key.
    class-data:error      type ref to i_oi_error,
               t_errors   type standard table of ref to i_oi_error with non-unique default key,
               cl_control type ref to i_oi_container_control.   "Office Dokument

endclass.                    "lcl_perform DEFINITION


start-of-selection.
  lcl_perform=>collect_reports( ).
  lcl_perform=>setup_objects( ).

end-of-selection.

  write '.'.  " Force output


*----------------------------------------------------------------------*
*       CLASS lcl_perform IMPLEMENTATION
*----------------------------------------------------------------------*
class lcl_perform implementation.
  method setup_objects.
    data: lo_split     type ref to cl_gui_splitter_container,
          lo_container type ref to cl_gui_container.

    data: it_fieldcat type lvc_t_fcat,
          is_layout   type lvc_s_layo,
          is_variant  type disvariant.
    field-symbols: <fc> like line of it_fieldcat.


    create object lo_split
      exporting
        parent                  = cl_gui_container=>screen0
        rows                    = 1
        columns                 = 3
        no_autodef_progid_dynnr = 'X'.
    lo_split->set_column_width(  exporting id                = 1
                                           width             = 20 ).
    lo_split->set_column_width(  exporting id                = 2
                                           width             = 40 ).

* Left:   List of reports
    lo_container = lo_split->get_container( row       = 1
                                            column    = 1 ).

    create object lo_grid
      exporting
        i_parent = lo_container.
    set handler lcl_perform=>handle_nav for lo_grid.

    is_variant-report = sy-repid.
    is_variant-handle = '0001'.

    is_layout-cwidth_opt = 'X'.

    append initial line to it_fieldcat assigning <fc>.
    <fc>-fieldname   = 'PROGNAME'.
    <fc>-tabname     = 'REPOSRC'.

    append initial line to it_fieldcat assigning <fc>.
    <fc>-fieldname   = 'SORT'.
    <fc>-ref_field   = 'PROGNAME'.
    <fc>-ref_table   = 'REPOSRC'.
    <fc>-tech        = abap_true. "No need to display this help field

    append initial line to it_fieldcat assigning <fc>.
    <fc>-fieldname   = 'DESCRIPTION'.
    <fc>-ref_field   = 'REPTI'.
    <fc>-ref_table   = 'RS38M'.

    lo_grid->set_table_for_first_display( exporting
                                            is_variant                    = is_variant
                                            i_save                        = 'A'
                                            is_layout                     = is_layout
                                          changing
                                            it_outtab                     = t_reports
                                            it_fieldcatalog               = it_fieldcat
                                          exceptions
                                            invalid_parameter_combination = 1
                                            program_error                 = 2
                                            too_many_lines                = 3
                                            others                        = 4 ).

* Middle: Text with coding
    lo_container = lo_split->get_container( row       = 1
                                            column    = 2 ).
    create object lo_text
      exporting
        parent = lo_container.
    lo_text->set_readonly_mode( cl_gui_textedit=>true ).
    lo_text->set_font_fixed( ).



* right:  DemoOutput
    lo_container = lo_split->get_container( row       = 1
                                            column    = 3 ).

    c_oi_container_control_creator=>get_container_control( importing control = cl_control
                                                                     error   = error ).
    append error to t_errors.

    cl_control->init_control( exporting  inplace_enabled     = 'X'
                                         no_flush            = 'X'
                                         r3_application_name = 'Demo Document Container'
                                         parent              = lo_container
                              importing  error               = error
                              exceptions others              = 2 ).
    append error to t_errors.

    cl_control->get_document_proxy( exporting document_type  = 'Excel.Sheet'                " EXCEL
                                              no_flush       = ' '
                                    importing document_proxy = cl_document
                                              error          = error ).
    append error to t_errors.
* Errorhandling should be inserted here


  endmethod.                    "setup_objects

  "collect_reports
  method collect_reports.
    field-symbols <report> like line of t_reports.
    data t_source type standard table of text255 with non-unique default key.
    data texts type standard table of textpool.
    data description type textpool.

* Get all demoreports
    select progname
      into corresponding fields of table t_reports
      from reposrc
      where progname like 'ZDEMO_EXCEL%'
        and progname <> sy-repid
        and subc     = '1'.

    loop at t_reports assigning <report>.

* Check if already switched to new outputoptions
      read report <report>-progname into t_source.
      if sy-subrc = 0.
        find 'INCLUDE zdemo_excel_outputopt_incl.' in table t_source ignoring case.
      endif.
      if sy-subrc <> 0.
        delete t_reports.
        continue.
      endif.


* Build half-numeric sort
      <report>-sort = <report>-progname.
      replace regex '(ZDEMO_EXCEL)(\d\d)\s*$' in <report>-sort with '$1\0$2'. "      REPLACE REGEX '(ZDEMO_EXCEL)([^][^])*$' IN <report>-sort WITH '$1$2'.REPLACE REGEX '(ZDEMO_EXCEL)([^][^])*$' IN <report>-sort WITH '$1$2'.REPLACE

      replace regex '(ZDEMO_EXCEL)(\d)\s*$'      in <report>-sort with '$1\0\0$2'.

* get report text
      read textpool <report>-progname into texts language sy-langu.
      read table texts into description with key id = 'R'.
      if sy-subrc > 0.
        "If not available in logon language, use english
        read textpool <report>-progname into texts language 'E'.
        read table texts into description with key id = 'R'.
      endif.
      "set report title
      <report>-description = description-entry.

    endloop.

    sort t_reports by sort progname.

  endmethod.  "collect_reports

  method handle_nav.
    constants: filename type text80 value 'ZABAP2XLSX_DEMO_SHOW.xlsx'.
    data: wa_report  like line of t_reports,
          t_source   type standard table of text255,
          t_rawdata  type solix_tab,
          wa_rawdata like line of t_rawdata,
          bytecount  type i,
          length     type i,
          add_selopt type flag.


    read table t_reports into wa_report index e_row-index.
    check sy-subrc = 0.

* Set new text into middle frame
    read report wa_report-progname into t_source.
    lo_text->set_text_as_r3table( exporting table = t_source ).


* Unload old xls-file
    cl_document->close_document( ).

* Get the demo
* If additional parameters found on selection screen, start via selection screen , otherwise start w/o
    clear add_selopt.
    find 'PARAMETERS' in table t_source.
    if sy-subrc = 0.
      add_selopt = 'X'.
    else.
      find 'SELECT-OPTIONS' in table t_source.
      if sy-subrc = 0.
        add_selopt = 'X'.
      endif.
    endif.
    if add_selopt is initial.
      submit (wa_report-progname) and return             "#EC CI_SUBMIT
              with p_backfn = filename
              with rb_back  = 'X'
              with rb_down  = ' '
              with rb_send  = ' '
              with rb_show  = ' '.
    else.
      submit (wa_report-progname) via selection-screen and return "#EC CI_SUBMIT
              with p_backfn = filename
              with rb_back  = 'X'
              with rb_down  = ' '
              with rb_send  = ' '
              with rb_show  = ' '.
    endif.

    open dataset filename for input in binary mode.
    if sy-subrc = 0.
      do.
        clear wa_rawdata.
        read dataset filename into wa_rawdata length length.
        if sy-subrc <> 0.
          append wa_rawdata to t_rawdata.
          add length to bytecount.
          exit.
        endif.
        append wa_rawdata to t_rawdata.
        add length to bytecount.
      enddo.
      close dataset filename.
    endif.

    cl_control->get_document_proxy( exporting document_type  = 'Excel.Sheet'                " EXCEL
                                              no_flush       = ' '
                                    importing document_proxy = cl_document
                                              error          = error ).

    cl_document->open_document_from_table( exporting document_size    = bytecount
                                                     document_table   = t_rawdata
                                                     open_inplace     = 'X' ).

  endmethod.                    "handle_nav

endclass.                    "lcl_perform IMPLEMENTATION
