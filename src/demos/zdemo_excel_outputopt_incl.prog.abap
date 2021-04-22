*&---------------------------------------------------------------------*
*&  Include           ZDEMO_EXCEL_OUTPUTOPT_INCL
*&---------------------------------------------------------------------*
class lcl_output definition create private.
  public section.
    class-methods:
      output         importing cl_excel            type ref to zcl_excel
                               iv_writerclass_name type clike optional,
      f4_path        returning value(selected_folder) type string,
      parametertexts.

  private section.
    methods:
      download_frontend,
      download_backend,
      display_online,
      send_email.

    data: xdata     type xstring,             " Will be used for sending as email
          t_rawdata type solix_tab,           " Will be used for downloading or open directly
          bytecount type i.                   " Will be used for downloading or open directly
endclass.                    "lcl_output DEFINITION


selection-screen begin of block bl1 with frame title txt_bl1.
parameters: rb_down radiobutton group rb1 user-command space.

parameters: rb_back radiobutton group rb1.

parameters: rb_show radiobutton group rb1 default 'X' .

parameters: rb_send radiobutton group rb1.

parameters: p_path  type string lower case modif id pat.
parameters: p_email type string lower case modif id ema obligatory default 'insert_your_emailadress@here'.
parameters: p_backfn type text40 no-display.
selection-screen end of block bl1.


at selection-screen output.
  loop at screen.

    if rb_down is initial and screen-group1 = 'PAT'.
      screen-input = 0.
      screen-invisible = 1.
    endif.

    if rb_send is initial and screen-group1 = 'EMA'.
      screen-input = 0.
      screen-invisible = 1.
    endif.

    modify screen.

  endloop.

initialization.
  if sy-batch is initial.
    cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = p_path ).
    cl_gui_cfw=>flush( ).
  endif.
  lcl_output=>parametertexts( ).  " If started in language w/o textelements translated set defaults
  sy-title = gc_save_file_name.
  txt_bl1 = 'Output options'(bl1).
  p_backfn = gc_save_file_name.  " Use as default if nothing else is supplied by submit

at selection-screen on value-request for p_path.
  p_path = lcl_output=>f4_path( ).


*----------------------------------------------------------------------*
*       CLASS lcl_output IMPLEMENTATION
*----------------------------------------------------------------------*
class lcl_output implementation.
  method output.

    data: cl_output type ref to lcl_output,
          cl_writer type ref to zif_excel_writer.

    if iv_writerclass_name is initial.
      create object cl_output.
      create object cl_writer type zcl_excel_writer_2007.
    else.
      create object cl_output.
      create object cl_writer type (iv_writerclass_name).
    endif.
    cl_output->xdata = cl_writer->write_file( cl_excel ).

* After 6.40 via cl_bcs_convert
    cl_output->t_rawdata = cl_bcs_convert=>xstring_to_solix( iv_xstring  = cl_output->xdata ).
    cl_output->bytecount = xstrlen( cl_output->xdata ).

* before 6.40
*  CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
*    EXPORTING
*      buffer        = cl_output->xdata
*    IMPORTING
*      output_length = cl_output->bytecount
*    TABLES
*      binary_tab    = cl_output->t_rawdata.

    case 'X'.
      when rb_down.
        if sy-batch is initial.
          cl_output->download_frontend( ).
        else.
          message e802(zabap2xlsx).
        endif.

      when rb_back.
        cl_output->download_backend( ).

      when rb_show.
        if sy-batch is initial.
          cl_output->display_online( ).
        else.
          message e803(zabap2xlsx).
        endif.

      when rb_send.
        cl_output->send_email( ).

    endcase.
  endmethod.                    "output

  method f4_path.
    data: new_path      type string,
          repid         type syrepid,
          dynnr         type sydynnr,
          lt_dynpfields type table of dynpread,
          ls_dynpfields like line of lt_dynpfields.

* Get current value
    dynnr = sy-dynnr.
    repid = sy-repid.
    ls_dynpfields-fieldname = 'P_PATH'.
    append ls_dynpfields to lt_dynpfields.

    call function 'DYNP_VALUES_READ'
      exporting
        dyname               = repid
        dynumb               = dynnr
      tables
        dynpfields           = lt_dynpfields
      exceptions
        invalid_abapworkarea = 1
        invalid_dynprofield  = 2
        invalid_dynproname   = 3
        invalid_dynpronummer = 4
        invalid_request      = 5
        no_fielddescription  = 6
        invalid_parameter    = 7
        undefind_error       = 8
        double_conversion    = 9
        stepl_not_found      = 10
        others               = 11.
    if sy-subrc <> 0.
      message id sy-msgid type sy-msgty number sy-msgno
              with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      exit.
    endif.

    read table lt_dynpfields into ls_dynpfields index 1.

    new_path = ls_dynpfields-fieldvalue.
    selected_folder = new_path.

    cl_gui_frontend_services=>directory_browse(
      exporting
        window_title         = 'Select path to download EXCEL-file'
        initial_folder       = new_path
      changing
        selected_folder      = new_path
      exceptions
        cntl_error           = 1
        error_no_gui         = 2
        not_supported_by_gui = 3
        others               = 4
           ).
    cl_gui_cfw=>flush( ).
    check new_path is not initial.
    selected_folder = new_path.

  endmethod.                                                "f4_path

  method parametertexts.
* If started in language w/o textelements translated set defaults
* Furthermore I don't have to change the selectiontexts of all demoreports.
    define default_parametertext.
      if %_&1_%_app_%-text = '&1' or
         %_&1_%_app_%-text is initial.
        %_&1_%_app_%-text = &2.
      endif.
    end-of-definition.

    default_parametertext:  rb_down  'Save to frontend',
                            rb_back  'Save to backend',
                            rb_show  'Direct display',
                            rb_send  'Send via email',

                            p_path   'Frontend-path to download to',
                            p_email  'Email to send xlsx to'.

  endmethod.                    "parametertexts

  method: download_frontend.
    data: filename type string.
* I don't like p_path here - but for this include it's ok
    filename = p_path.
* Add trailing "\" or "/"
    if filename ca '/'.
      replace regex '([^/])\s*$' in filename with '$1/' .
    else.
      replace regex '([^\\])\s*$' in filename with '$1\\'.
    endif.

    concatenate filename gc_save_file_name into filename.
* Get trailing blank
    cl_gui_frontend_services=>gui_download( exporting bin_filesize = bytecount
                                                      filename     = filename
                                                      filetype     = 'BIN'
                                             changing data_tab     = t_rawdata ).
  endmethod.                    "download_frontend

  method download_backend.
    data: bytes_remain type i.
    field-symbols: <rawdata> like line of t_rawdata.

    open dataset p_backfn for output in binary mode.
    check sy-subrc = 0.

    bytes_remain = bytecount.

    loop at t_rawdata assigning <rawdata>.

      at last.
        check bytes_remain >= 0.
        transfer <rawdata> to p_backfn length bytes_remain.
        exit.
      endat.

      transfer <rawdata> to p_backfn.
      subtract 255 from bytes_remain.  " Solix has length 255

    endloop.

    close dataset p_backfn.

    if sy-repid <> sy-cprog and sy-cprog is not initial.  " no need to display anything if download was selected and report was called for demo purposes
      leave program.
    else.
      message 'Data transferred to default backend directory' type 'S'.
    endif.
  endmethod.                    "download_backend

  method display_online.
    data:error       type ref to i_oi_error,
         t_errors    type standard table of ref to i_oi_error with non-unique default key,
         cl_control  type ref to i_oi_container_control, "OIContainerCtrl
         cl_document type ref to i_oi_document_proxy.   "Office Dokument

    c_oi_container_control_creator=>get_container_control( importing control = cl_control
                                                                     error   = error ).
    append error to t_errors.

    cl_control->init_control( exporting  inplace_enabled     = 'X'
                                         no_flush            = 'X'
                                         r3_application_name = 'Demo Document Container'
                                         parent              = cl_gui_container=>screen0
                              importing  error               = error
                              exceptions others              = 2 ).
    append error to t_errors.

    cl_control->get_document_proxy( exporting document_type  = 'Excel.Sheet'                " EXCEL
                                              no_flush       = ' '
                                    importing document_proxy = cl_document
                                              error          = error ).
    append error to t_errors.
* Errorhandling should be inserted here

    cl_document->open_document_from_table( exporting document_size    = bytecount
                                                     document_table   = t_rawdata
                                                     open_inplace     = 'X' ).

    write: '.'.  " To create an output.  That way screen0 will exist
  endmethod.                    "display_online

  method send_email.
* Needed to send emails
    data: bcs_exception        type ref to cx_bcs,
          errortext            type string,
          cl_send_request      type ref to cl_bcs,
          cl_document          type ref to cl_document_bcs,
          cl_recipient         type ref to if_recipient_bcs,
          cl_sender            type ref to cl_cam_address_bcs,
          t_attachment_header  type soli_tab,
          wa_attachment_header like line of t_attachment_header,
          attachment_subject   type sood-objdes,

          sood_bytecount       type sood-objlen,
          mail_title           type so_obj_des,
          t_mailtext           type soli_tab,
          wa_mailtext          like line of t_mailtext,
          send_to              type adr6-smtp_addr,
          sent                 type os_boolean.


    mail_title     = 'Mail title'.
    wa_mailtext    = 'Mailtext'.
    append wa_mailtext to t_mailtext.

    try.
* Create send request
        cl_send_request = cl_bcs=>create_persistent( ).
* Create new document with mailtitle and mailtextg
        cl_document = cl_document_bcs=>create_document( i_type    = 'RAW' "#EC NOTEXT
                                                        i_text    = t_mailtext
                                                        i_subject = mail_title ).
* Add attachment to document
* since the new excelfiles have an 4-character extension .xlsx but the attachment-type only holds 3 charactes .xls,
* we have to specify the real filename via attachment header
* Use attachment_type xls to have SAP display attachment with the excel-icon
        attachment_subject  = gc_save_file_name.
        concatenate '&SO_FILENAME=' attachment_subject into wa_attachment_header.
        append wa_attachment_header to t_attachment_header.
* Attachment
        sood_bytecount = bytecount.  " next method expects sood_bytecount instead of any positive integer *sigh*
        cl_document->add_attachment(  i_attachment_type    = 'XLS' "#EC NOTEXT
                                      i_attachment_subject = attachment_subject
                                      i_attachment_size    = sood_bytecount
                                      i_att_content_hex    = t_rawdata
                                      i_attachment_header  = t_attachment_header ).

* add document to send request
        cl_send_request->set_document( cl_document ).

* set sender in case if no own email is availabe
*        cl_sender  = cl_cam_address_bcs=>create_internet_address( 'sender@sender.sender' ).
*        cl_send_request->set_sender( cl_sender ).

* add recipient(s) - here only 1 will be needed
        send_to = p_email.
        if send_to is initial.
          send_to = 'no_email@no_email.no_email'.  " Place into SOST in any case for demonstration purposes
        endif.
        cl_recipient = cl_cam_address_bcs=>create_internet_address( send_to ).
        cl_send_request->add_recipient( cl_recipient ).

* Und abschicken
        sent = cl_send_request->send( i_with_error_screen = 'X' ).

        commit work.

        if sent is initial.
          message i804(zabap2xlsx) with p_email.
        else.
          message s805(zabap2xlsx).
          message 'Document ready to be sent - Check SOST or SCOT' type 'I'.
        endif.

      catch cx_bcs into bcs_exception.
        errortext = bcs_exception->if_message~get_text( ).
        message errortext type 'I'.

    endtry.
  endmethod.                    "send_email


endclass.                    "lcl_output IMPLEMENTATION
