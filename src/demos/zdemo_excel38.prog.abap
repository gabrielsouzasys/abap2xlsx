report zdemo_excel38.


data: lo_excel     type ref to zcl_excel,
      lo_worksheet type ref to zcl_excel_worksheet,
      lo_column    type ref to zcl_excel_column,
      lo_drawing   type ref to zcl_excel_drawing.

types: begin of gty_icon,
*         name      TYPE icon_name, "Fix #228
         name  type iconname,   "Fix #228
         objid type w3objid,
       end of gty_icon,
       gtyt_icon type standard table of gty_icon with non-unique default key.

data: lt_icon       type gtyt_icon,
      lv_row        type sytabix,
      ls_wwwdatatab type wwwdatatab,
      lt_mimedata   type standard table of w3mime with non-unique default key,
      lv_xstring    type xstring.

field-symbols: <icon>     like line of lt_icon,
               <mimedata> like line of lt_mimedata.

constants: gc_save_file_name type string value '38_SAP-Icons.xlsx'.
include zdemo_excel_outputopt_incl.


tables: icon.
select-options: s_icon for icon-name default 'ICON_LED_*' option cp.

start-of-selection.
  " Creates active sheet
  create object lo_excel.

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = 'Demo Icons' ).
  lo_column = lo_worksheet->get_column( ip_column = 'A' ).
  lo_column->set_auto_size( 'X' ).
  lo_column = lo_worksheet->get_column( ip_column = 'B' ).
  lo_column->set_auto_size( 'X' ).

* Get all icons
  select name
    into table lt_icon
    from icon
    where name in s_icon
    order by name.
  loop at lt_icon assigning <icon>.

    lv_row = sy-tabix.
*--------------------------------------------------------------------*
* Set name of icon
*--------------------------------------------------------------------*
    lo_worksheet->set_cell( ip_row = lv_row
                            ip_column = 'A'
                            ip_value = <icon>-name ).
*--------------------------------------------------------------------*
* Check whether the mime-repository holds some icondata for us
*--------------------------------------------------------------------*

* Get key
    select single objid
      into <icon>-objid
      from wwwdata
      where text = <icon>-name.
    check sy-subrc = 0.  " :o(
    lo_worksheet->set_cell( ip_row = lv_row
                            ip_column = 'B'
                            ip_value = <icon>-objid ).

* Load mimedata
    clear lt_mimedata.
    clear ls_wwwdatatab.
    ls_wwwdatatab-relid = 'MI' .
    ls_wwwdatatab-objid = <icon>-objid.
    call function 'WWWDATA_IMPORT'
      exporting
        key               = ls_wwwdatatab
      tables
        mime              = lt_mimedata
      exceptions
        wrong_object_type = 1
        import_error      = 2
        others            = 3.
    check sy-subrc = 0.  " :o(

    lo_drawing = lo_excel->add_new_drawing( ).
    lo_drawing->set_position( ip_from_row = lv_row
                              ip_from_col = 'C' ).
    clear lv_xstring.
    loop at lt_mimedata assigning <mimedata>.
      concatenate lv_xstring <mimedata>-line into lv_xstring in byte mode.
    endloop.

    lo_drawing->set_media( ip_media      = lv_xstring
                           ip_media_type = zcl_excel_drawing=>c_media_type_jpg
                           ip_width      = 16
                           ip_height     = 14  ).
    lo_worksheet->add_drawing( lo_drawing ).

  endloop.

*** Create output
  lcl_output=>output( lo_excel ).
