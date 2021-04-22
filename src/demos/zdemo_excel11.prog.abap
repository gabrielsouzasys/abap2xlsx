*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL11
*& Export Organisation and Contact Persons using ABAP2XLSX
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel11.

type-pools: abap.

data: central_search     type bapibus1006_central_search,
      addressdata_search type bapibus1006_addr_search,
      others_search      type bapibus1006_other_data.
data: searchresult type table of bapibus1006_bp_addr,
      return       type table of bapiret2.
data: lines type i.
field-symbols: <searchresult_line> like line of searchresult.
data: centraldata             type bapibus1006_central,
      centraldataperson       type bapibus1006_central_person,
      centraldataorganization type bapibus1006_central_organ.
data: addressdata type bapibus1006_address.
data: relationships type table of bapibus1006_relations.
field-symbols: <relationship> like line of relationships.
data: relationship_centraldata type bapibus1006002_central.
data: relationship_addresses type table of bapibus1006002_addresses.
field-symbols: <relationship_address> like line of relationship_addresses.

data: lt_download type table of zexcel_s_org_rel.
field-symbols: <download> like line of lt_download.

constants: gc_save_file_name type string value '11_Export_Org_and_Contact.xlsx'.
include zdemo_excel_outputopt_incl.


parameters: md type flag radiobutton group act.

selection-screen begin of block a with frame title text-00a.
parameters: partnerc type bu_type   default 2, " Organizations
            postlcod type ad_pstcd1 default '8334*',
            country  type land1     default 'DE',
            maxsel   type bu_maxsel default 100.
selection-screen end of block a.

parameters: rel type flag radiobutton group act default 'X'.

selection-screen begin of block b with frame title text-00b.
parameters: reltyp  type bu_reltyp default 'BUR011',
            partner type bu_partner default '191'.
selection-screen end of block b.

start-of-selection.
  if md = abap_true.
    " Read all Companies by Master Data
    central_search-partnercategory = partnerc.
    addressdata_search-postl_cod1  = postlcod.
    addressdata_search-country     = country.
    others_search-maxsel           = maxsel.
    others_search-no_search_for_contactperson = 'X'.

    call function 'BAPI_BUPA_SEARCH_2'
      exporting
        centraldata  = central_search
        addressdata  = addressdata_search
        others       = others_search
      tables
        searchresult = searchresult
        return       = return.

    sort searchresult by partner.
    delete adjacent duplicates from searchresult comparing partner.
  elseif rel = abap_true.
    " Read by Relationship
    select but050~partner1 as partner from but050
      inner join but000 on but000~partner = but050~partner1 and but000~type = '2'
      into corresponding fields of table searchresult
      where but050~partner2 = partner
        and but050~reltyp   = reltyp.
  endif.

  describe table searchresult lines lines.
  write: / 'Number of search results: ', lines.

  loop at searchresult assigning <searchresult_line>.
    " Read Details of Organization
    call function 'BAPI_BUPA_CENTRAL_GETDETAIL'
      exporting
        businesspartner         = <searchresult_line>-partner
      importing
        centraldataorganization = centraldataorganization.
    " Read Standard Address of Organization
    call function 'BAPI_BUPA_ADDRESS_GETDETAIL'
      exporting
        businesspartner = <searchresult_line>-partner
      importing
        addressdata     = addressdata.

    " Add Organization to Download
    append initial line to lt_download assigning <download>.
    " Fill Organization Partner Numbers
    call function 'BAPI_BUPA_GET_NUMBERS'
      exporting
        businesspartner        = <searchresult_line>-partner
      importing
        businesspartnerout     = <download>-org_number
        businesspartnerguidout = <download>-org_guid.

    move-corresponding centraldataorganization to <download>.
    move-corresponding addressdata to <download>.
    clear: addressdata.

    " Read all Relationships
    clear: relationships.
    call function 'BAPI_BUPA_RELATIONSHIPS_GET'
      exporting
        businesspartner = <searchresult_line>-partner
      tables
        relationships   = relationships.
    delete relationships where relationshipcategory <> 'BUR001'.
    loop at relationships assigning <relationship>.
      " Read details of Contact person
      call function 'BAPI_BUPA_CENTRAL_GETDETAIL'
        exporting
          businesspartner   = <relationship>-partner2
        importing
          centraldata       = centraldata
          centraldataperson = centraldataperson.
      " Read details of the Relationship
      call function 'BAPI_BUPR_CONTP_GETDETAIL'
        exporting
          businesspartner = <relationship>-partner1
          contactperson   = <relationship>-partner2
        importing
          centraldata     = relationship_centraldata.
      " Read relationship address
      clear: relationship_addresses.

      call function 'BAPI_BUPR_CONTP_ADDRESSES_GET'
        exporting
          businesspartner = <relationship>-partner1
          contactperson   = <relationship>-partner2
        tables
          addresses       = relationship_addresses.

      read table relationship_addresses
        assigning <relationship_address>
        with key standardaddress = 'X'.

      if sy-subrc = 0.
        " Read Relationship Address
        clear addressdata.
        call function 'BAPI_BUPA_ADDRESS_GETDETAIL'
          exporting
            businesspartner = <searchresult_line>-partner
            addressguid     = <relationship_address>-addressguid
          importing
            addressdata     = addressdata.

        append initial line to lt_download assigning <download>.
        call function 'BAPI_BUPA_GET_NUMBERS'
          exporting
            businesspartner        = <relationship>-partner1
          importing
            businesspartnerout     = <download>-org_number
            businesspartnerguidout = <download>-org_guid.

        call function 'BAPI_BUPA_GET_NUMBERS'
          exporting
            businesspartner        = <relationship>-partner2
          importing
            businesspartnerout     = <download>-contpers_number
            businesspartnerguidout = <download>-contpers_guid.

        move-corresponding centraldataorganization to <download>.
        move-corresponding addressdata to <download>.
        move-corresponding centraldataperson to <download>.
        move-corresponding relationship_centraldata to <download>.

        write: / <relationship>-partner1, <relationship>-partner2.
        write: centraldataorganization-name1(20), centraldataorganization-name2(10).
        write: centraldataperson-firstname(15), centraldataperson-lastname(15).
        write: addressdata-street(25), addressdata-house_no,
               addressdata-postl_cod1, addressdata-city(25).
      endif.
    endloop.

  endloop.

  data: lo_excel       type ref to zcl_excel,
        lo_worksheet   type ref to zcl_excel_worksheet,
        lo_style_body  type ref to zcl_excel_style,
        lo_border_dark type ref to zcl_excel_style_border,
        lo_column      type ref to zcl_excel_column,
        lo_row         type ref to zcl_excel_row.

  data: lv_style_body_even_guid type zexcel_cell_style,
        lv_style_body_green     type zexcel_cell_style.

  data: row type zexcel_cell_row.

  data: lt_field_catalog  type zexcel_t_fieldcatalog,
        ls_table_settings type zexcel_s_table_settings.

  data: column       type zexcel_cell_column,
        column_alpha type zexcel_cell_column_alpha,
        value        type zexcel_cell_value.

  field-symbols: <fs_field_catalog> type zexcel_s_fieldcatalog.

  " Creates active sheet
  create object lo_excel.

  " Create border object
  create object lo_border_dark.
  lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
  lo_border_dark->border_style = zcl_excel_style_border=>c_border_thin.
  "Create style with border even
  lo_style_body                       = lo_excel->add_new_style( ).
  lo_style_body->fill->fgcolor-rgb    = zcl_excel_style_color=>c_yellow.
  lo_style_body->borders->allborders  = lo_border_dark.
  lv_style_body_even_guid             = lo_style_body->get_guid( ).
  "Create style with border and green fill
  lo_style_body                       = lo_excel->add_new_style( ).
  lo_style_body->fill->fgcolor-rgb    = zcl_excel_style_color=>c_green.
  lo_style_body->borders->allborders  = lo_border_dark.
  lv_style_body_green                 = lo_style_body->get_guid( ).

  " Get active sheet
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'Internal table' ).

  lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = lt_download ).

  loop at lt_field_catalog assigning <fs_field_catalog>.
    case <fs_field_catalog>-fieldname.
      when 'ORG_NUMBER'.
        <fs_field_catalog>-position   = 1.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'CONTPERS_NUMBER'.
        <fs_field_catalog>-position   = 2.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'NAME1'.
        <fs_field_catalog>-position   = 3.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'NAME2'.
        <fs_field_catalog>-position   = 4.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'STREET'.
        <fs_field_catalog>-position   = 5.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'HOUSE_NO'.
        <fs_field_catalog>-position   = 6.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'POSTL_COD1'.
        <fs_field_catalog>-position   = 7.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'CITY'.
        <fs_field_catalog>-position   = 8.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'COUNTRYISO'.
        <fs_field_catalog>-position   = 9.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'FIRSTNAME'.
        <fs_field_catalog>-position   = 10.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'LASTNAME'.
        <fs_field_catalog>-position   = 11.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'FUNCTIONNAME'.
        <fs_field_catalog>-position   = 12.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'DEPARTMENTNAME'.
        <fs_field_catalog>-position   = 13.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'TEL1_NUMBR'.
        <fs_field_catalog>-position   = 14.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'TEL1_EXT'.
        <fs_field_catalog>-position   = 15.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'FAX_NUMBER'.
        <fs_field_catalog>-position   = 16.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'FAX_EXTENS'.
        <fs_field_catalog>-position   = 17.
        <fs_field_catalog>-dynpfld    = abap_true.
      when 'E_MAIL'.
        <fs_field_catalog>-position   = 18.
        <fs_field_catalog>-dynpfld    = abap_true.
      when others.
        <fs_field_catalog>-dynpfld = abap_false.
    endcase.
  endloop.

  ls_table_settings-top_left_column = 'A'.
  ls_table_settings-top_left_row    = '2'.
  ls_table_settings-table_style  = zcl_excel_table=>builtinstyle_medium5.

  lo_worksheet->bind_table( ip_table          = lt_download
                            is_table_settings = ls_table_settings
                            it_field_catalog  = lt_field_catalog ).
  loop at lt_download assigning <download>.
    row = sy-tabix + 2.
    if    not <download>-org_number is initial
      and <download>-contpers_number is initial.
      " Mark fields of Organization which can be changed green
      lo_worksheet->set_cell_style(
          ip_column = 'C'
          ip_row    = row
          ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'D'
          ip_row    = row
          ip_style  = lv_style_body_green
      ).
*        CATCH zcx_excel.    " Exceptions for ABAP2XLSX
    elseif not <download>-org_number is initial
       and not <download>-contpers_number is initial.
      " Mark fields of Relationship which can be changed green
      lo_worksheet->set_cell_style(
          ip_column = 'L' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'M' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'N' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'O' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'P' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'Q' ip_row    = row ip_style  = lv_style_body_green
      ).
      lo_worksheet->set_cell_style(
          ip_column = 'R' ip_row    = row ip_style  = lv_style_body_green
      ).
    endif.
  endloop.
  " Add Fieldnames in first row and hide the row
  loop at lt_field_catalog assigning <fs_field_catalog>
    where position <> '' and dynpfld = abap_true.
    column = <fs_field_catalog>-position.
    column_alpha = zcl_excel_common=>convert_column2alpha( column ).
    value = <fs_field_catalog>-fieldname.
    lo_worksheet->set_cell( ip_column = column_alpha
                            ip_row    = 1
                            ip_value  = value
                            ip_style  = lv_style_body_even_guid ).
  endloop.
  " Hide first row
  lo_row = lo_worksheet->get_row( 1 ).
  lo_row->set_visible( abap_false ).

  data: highest_column type zexcel_cell_column,
        count          type int4,
        col_alpha      type zexcel_cell_column_alpha.

  highest_column = lo_worksheet->get_highest_column( ).
  count = 1.
  while count <= highest_column.
    col_alpha = zcl_excel_common=>convert_column2alpha( ip_column = count ).
    lo_column = lo_worksheet->get_column( ip_column = col_alpha ).
    lo_column->set_auto_size( ip_auto_size = abap_true ).
    count = count + 1.
  endwhile.

*** Create output
  lcl_output=>output( lo_excel ).
