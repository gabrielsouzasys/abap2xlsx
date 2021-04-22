report zdemo_excel41.

constants: gc_save_file_name type string value 'ABAP2XLSX Inheritance.xlsx'.

*--------------------------------------------------------------------*
* Demo inheritance ZCL_EXCEL1
* Variation of ZCL_EXCEL that creates numerous sheets
*--------------------------------------------------------------------*
class lcl_my_zcl_excel1 definition inheriting from zcl_excel.
  public section.
    methods: constructor importing iv_sheetcount type i default 5
                         raising   zcx_excel.
endclass.

class lcl_my_zcl_excel1 implementation.
  method constructor.
    data: lv_sheets_to_create type i.
    super->constructor( ).
    lv_sheets_to_create = iv_sheetcount - 1. " one gets created by standard class
    do lv_sheets_to_create times.
      try.
          me->add_new_worksheet( ).
        catch zcx_excel.
      endtry.
    enddo.
    me->set_active_sheet_index( 1 ).

  endmethod.
endclass.

*--------------------------------------------------------------------*
* Demo inheritance ZCL_EXCEL_WORKSHEET
* Variation of ZCL_EXCEL_WORKSHEET ( and ZCL_EXCEL that calls the new type of worksheet )
* that sets a fixed title
*--------------------------------------------------------------------*
class lcl_my_zcl_excel2 definition inheriting from zcl_excel.
  public section.
    methods: constructor raising zcx_excel.
endclass.

class lcl_my_zcl_excel_worksheet definition inheriting from zcl_excel_worksheet.
  public section.
    methods: constructor importing ip_excel type ref to zcl_excel
                                   ip_title type zexcel_sheet_title optional  " Will be ignored - keep parameter for demonstration purpose
                         raising   zcx_excel.
endclass.

class lcl_my_zcl_excel2 implementation.
  method constructor.

    data: lo_worksheet type ref to zcl_excel_worksheet.

    super->constructor( ).

* To use own worksheet we have to remove the standard worksheet
    lo_worksheet = get_active_worksheet( ).
    me->worksheets->remove( lo_worksheet ).
* and replace it with own version
    create object lo_worksheet type lcl_my_zcl_excel_worksheet
      exporting
        ip_excel = me
        ip_title = 'This title will be ignored'.
    me->worksheets->add( lo_worksheet ).

  endmethod.
endclass.

class lcl_my_zcl_excel_worksheet implementation.
  method constructor.
    super->constructor( ip_excel = ip_excel
                        ip_title = 'Inherited Worksheet' ).

  endmethod.
endclass.

data: go_excel1 type ref to lcl_my_zcl_excel1.
data: go_excel2 type ref to lcl_my_zcl_excel2.


selection-screen begin of block bli with frame title text-bli.
parameters: rbi_1 radiobutton group rbi default 'X' , " Simple inheritance
            rbi_2 radiobutton group rbi.
selection-screen end of block bli.

include zdemo_excel_outputopt_incl.

end-of-selection.

  case 'X'.

    when rbi_1.                                       " Simple inheritance of zcl_excel, object created directly
      create object go_excel1
        exporting
          iv_sheetcount = 5.
      lcl_output=>output( go_excel1 ).

    when rbi_2.                                       " Inheritance of zcl_excel_worksheet, inheritance of zcl_excel needed to allow this
      create object go_excel2.
      lcl_output=>output( go_excel2 ).


  endcase.
