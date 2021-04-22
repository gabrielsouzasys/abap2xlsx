report zdemo_excel37.

type-pools: vrm.

data: excel                 type ref to zcl_excel,
      reader                type ref to zif_excel_reader,
      go_error              type ref to cx_root,
      gv_memid_gr8          type text255,
      gv_message            type string,
      lv_extension          type string,
      gv_error_program_name type syrepid,
      gv_error_include_name type syrepid,
      gv_error_line         type i.

data: gc_save_file_name type string value '37- Read template and output.&'.

selection-screen begin of block blx with frame.
parameters: p_upfile type string lower case memory id gr8.
selection-screen end of block blx.

include zdemo_excel_outputopt_incl.

selection-screen begin of block cls with frame title text-cls.
parameters: lb_read   type seoclsname as listbox visible length 40 lower case obligatory default 'Autodetect'(001).
parameters: lb_write  type seoclsname as listbox visible length 40 lower case obligatory default 'Autodetect'(001).
selection-screen end of block cls.

selection-screen begin of block bl_err with frame title text-err.
parameters: cb_errl as checkbox default 'X'.
selection-screen begin of line.
parameters: cb_dump as checkbox default space.
selection-screen comment (60) cmt_dump for field cb_dump.
selection-screen end of line.
selection-screen end of block bl_err.

initialization.
  perform setup_listboxes.
  cmt_dump = text-dum.
  get parameter id 'GR8' field gv_memid_gr8.
  p_upfile = gv_memid_gr8.

  if p_upfile is initial.
    p_upfile = 'c:\temp\whatever.xlsx'.
  endif.

at selection-screen on value-request for p_upfile.
  perform f4_p_upfile changing p_upfile.


start-of-selection.
  if cb_dump is initial.
    try.
        perform read_template.
        perform write_template.
*** Create output
      catch cx_root into go_error.
        message 'Error reading excelfile' type 'I'.
        gv_message = go_error->get_text( ).
        if cb_errl = ' '.
          if gv_message is not initial.
            message gv_message type 'I'.
          endif.
        else.
          go_error->get_source_position( importing program_name = gv_error_program_name
                                                   include_name = gv_error_include_name
                                                   source_line  = gv_error_line         ).
          write:/ 'Errormessage:'       ,gv_message.
          write:/ 'Errorposition:',
                at /10 'Program:'       ,gv_error_program_name,
                at /10 'include_name:'  ,gv_error_include_name,
                at /10 'source_line:'   ,gv_error_line.
        endif.
    endtry.
  else.  " This will dump if an error occurs.  In some cases the information given in cx_root is not helpful - this will show exactly where the problem is
    perform read_template.
    perform write_template.
  endif.



*&---------------------------------------------------------------------*
*&      Form  F4_P_UPFILE
*&---------------------------------------------------------------------*
form f4_p_upfile  changing p_upfile type string.

  data: lv_repid       type syrepid,
        lt_fields      type dynpread_tabtype,
        ls_field       like line of lt_fields,
        lt_files       type filetable,
        lv_file_filter type string.

  lv_repid = sy-repid.

  call function 'DYNP_VALUES_READ'
    exporting
      dyname               = lv_repid
      dynumb               = '1000'
      request              = 'A'
    tables
      dynpfields           = lt_fields
    exceptions
      invalid_abapworkarea = 01
      invalid_dynprofield  = 02
      invalid_dynproname   = 03
      invalid_dynpronummer = 04
      invalid_request      = 05
      no_fielddescription  = 06
      undefind_error       = 07.
  read table lt_fields into ls_field with key fieldname = 'P_UPFILE'.
  p_upfile = ls_field-fieldvalue.

  lv_file_filter = 'Excel Files (*.XLSX;*.XLSM)|*.XLSX;*.XLSM'.
  cl_gui_frontend_services=>file_open_dialog( exporting
                                                default_filename        = p_upfile
                                                file_filter             = lv_file_filter
                                              changing
                                                file_table              = lt_files
                                                rc                      = sy-tabix
                                              exceptions
                                                others                  = 1 ).
  read table lt_files index 1 into p_upfile.

endform.                    " F4_P_UPFILE


*&---------------------------------------------------------------------*
*&      Form  SETUP_LISTBOXES
*&---------------------------------------------------------------------*
form setup_listboxes .

  data: lv_id                   type vrm_id,
        lt_values               type vrm_values,
        lt_implementing_classes type seo_relkeys.

  field-symbols: <ls_implementing_class> like line of lt_implementing_classes,
                 <ls_value>              like line of lt_values.

*--------------------------------------------------------------------*
* Possible READER-Classes
*--------------------------------------------------------------------*
  lv_id = 'LB_READ'.
  append initial line to lt_values assigning <ls_value>.
  <ls_value>-key  = 'Autodetect'(001).
  <ls_value>-text = 'Autodetect'(001).


  perform get_implementing_classds using    'ZIF_EXCEL_READER'
                                   changing lt_implementing_classes.
  clear lt_values.
  loop at lt_implementing_classes assigning <ls_implementing_class>.

    append initial line to lt_values assigning <ls_value>.
    <ls_value>-key  = <ls_implementing_class>-clsname.
    <ls_value>-text = <ls_implementing_class>-clsname.

  endloop.

  call function 'VRM_SET_VALUES'
    exporting
      id              = lv_id
      values          = lt_values
    exceptions
      id_illegal_name = 1
      others          = 2.

*--------------------------------------------------------------------*
* Possible WRITER-Classes
*--------------------------------------------------------------------*
  lv_id = 'LB_WRITE'.
  append initial line to lt_values assigning <ls_value>.
  <ls_value>-key  = 'Autodetect'(001).
  <ls_value>-text = 'Autodetect'(001).


  perform get_implementing_classds using    'ZIF_EXCEL_WRITER'
                                   changing lt_implementing_classes.
  clear lt_values.
  loop at lt_implementing_classes assigning <ls_implementing_class>.

    append initial line to lt_values assigning <ls_value>.
    <ls_value>-key  = <ls_implementing_class>-clsname.
    <ls_value>-text = <ls_implementing_class>-clsname.

  endloop.

  call function 'VRM_SET_VALUES'
    exporting
      id              = lv_id
      values          = lt_values
    exceptions
      id_illegal_name = 1
      others          = 2.

endform.                    " SETUP_LISTBOXES


*&---------------------------------------------------------------------*
*&      Form  GET_IMPLEMENTING_CLASSDS
*&---------------------------------------------------------------------*
form get_implementing_classds  using    iv_interface_name       type clike
                               changing ct_implementing_classes type seo_relkeys.

  data: lo_oo_interface            type ref to cl_oo_interface,
        lo_oo_class                type ref to cl_oo_class,
        lt_implementing_subclasses type seo_relkeys.

  field-symbols: <ls_implementing_class> like line of ct_implementing_classes.

  try.
      lo_oo_interface ?= cl_oo_interface=>get_instance( iv_interface_name ).
    catch cx_class_not_existent.
      return.
  endtry.
  ct_implementing_classes = lo_oo_interface->get_implementing_classes( ).

  loop at ct_implementing_classes assigning <ls_implementing_class>.
    try.
        lo_oo_class ?= cl_oo_class=>get_instance( <ls_implementing_class>-clsname ).
        lt_implementing_subclasses = lo_oo_class->get_subclasses( ).
        append lines of lt_implementing_subclasses to ct_implementing_classes.
      catch cx_class_not_existent.
    endtry.
  endloop.


endform.                    " GET_IMPLEMENTING_CLASSDS


*&---------------------------------------------------------------------*
*&      Form  READ_TEMPLATE
*&---------------------------------------------------------------------*
form read_template raising zcx_excel .

  case lb_read.
    when 'Autodetect'(001).
      find regex '(\.xlsx|\.xlsm)\s*$' in p_upfile submatches lv_extension.
      translate lv_extension to upper case.
      case lv_extension.

        when '.XLSX'.
          create object reader type zcl_excel_reader_2007.
          excel = reader->load_file(  p_upfile ).
          "Use template for charts
          excel->use_template = abap_true.

        when '.XLSM'.
          create object reader type zcl_excel_reader_xlsm.
          excel = reader->load_file(  p_upfile ).
          "Use template for charts
          excel->use_template = abap_true.

        when others.
          message 'Unsupported filetype' type 'I'.
          return.

      endcase.

    when others.
      create object reader type (lb_read).
      excel = reader->load_file(  p_upfile ).
      "Use template for charts
      excel->use_template = abap_true.

  endcase.

endform.                    " READ_TEMPLATE


*&---------------------------------------------------------------------*
*&      Form  WRITE_TEMPLATE
*&---------------------------------------------------------------------*
form write_template  raising zcx_excel.

  case lb_write.

    when 'Autodetect'(001).
      find regex '(\.xlsx|\.xlsm)\s*$' in p_upfile submatches lv_extension.
      translate lv_extension to upper case.
      case lv_extension.

        when '.XLSX'.
          replace '&' in gc_save_file_name with 'xlsx'.  " Pass extension for standard writer
          lcl_output=>output( excel ).

        when '.XLSM'.
          replace '&' in gc_save_file_name with 'xlsm'.  " Pass extension for macro-writer
          lcl_output=>output( cl_excel            = excel
                              iv_writerclass_name = 'ZCL_EXCEL_WRITER_XLSM' ).

        when others.
          message 'Unsupported filetype' type 'I'.
          return.

      endcase.

    when others.
      lcl_output=>output( cl_excel            = excel
                          iv_writerclass_name = lb_write ).
  endcase.

endform.                    " WRITE_TEMPLATE
