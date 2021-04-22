*&---------------------------------------------------------------------*
*& Report  ZDEMO_EXCEL
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

report zdemo_excel.

data: lv_workdir type string,
      lv_upfile  type string.


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

  cl_gui_frontend_services=>get_file_separator( changing file_separator = sy-lisel ).
  concatenate p_path sy-lisel '01_HelloWorld.xlsx' into lv_upfile.

  submit zdemo_excel1 with rb_down = abap_true with rb_show = abap_false with  p_path     = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Hello world
  submit zdemo_excel2 with rb_down = abap_true with rb_show = abap_false with  p_path     = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Styles
  submit zdemo_excel3 with rb_down = abap_true with rb_show = abap_false with  p_path     = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: iTab binding
  submit zdemo_excel4 with rb_down = abap_true with rb_show = abap_false with  p_path     = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Multi sheets, page setup and sheet properties
  submit zdemo_excel5 with rb_down = abap_true with rb_show = abap_false with  p_path     = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Conditional formatting
  submit zdemo_excel6 with rb_down = abap_true with rb_show = abap_false with  p_path     = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Formulas
  submit zdemo_excel7 with rb_down = abap_true with rb_show = abap_false with  p_path     = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Conditional formatting
  submit zdemo_excel8 with rb_down = abap_true with rb_show = abap_false with  p_path     = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Ranges
  submit zdemo_excel9 with rb_down = abap_true with rb_show = abap_false with  p_path     = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Data validation
  submit zdemo_excel10 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Bind table with field catalog
  " zdemo_excel11 is not added because it has a selection screen and
  " you also need to have business partners maintained in transaction BP
  submit zdemo_excel12 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Column size
  submit zdemo_excel13 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Merge cell
  submit zdemo_excel14 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Alignment
  " zdemo_excel15 added at the end
  submit zdemo_excel16 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Drawing
  submit zdemo_excel17 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Lock sheet
  submit zdemo_excel18 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Lock workbook
  submit zdemo_excel19 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Set active sheet
  " zdemo_excel20 is not added because it uses ALV and cannot be processed (OLE2)
  submit zdemo_excel21 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Color Picker
  submit zdemo_excel22 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Bind table with field catalog & sheet style
  submit zdemo_excel23 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Multiple sheets with and w/o grid lines, print options
  submit zdemo_excel24 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Multiple sheets with different default date formats
  submit zdemo_excel25 and return. "#EC CI_SUBMIT abap2xlsx Demo: Create and xlsx on Application Server (could be executed in batch mode)
  " zdemo_excel26 is not added because it uses ALV and cannot be processed (Native)
  submit zdemo_excel27 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Conditional Formatting
  submit zdemo_excel28 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: CSV writer
  " SUBMIT zdemo_excel29 WITH rb_down = abap_true WITH rb_show = abap_false WITH  p_path  = p_path AND RETURN. "#EC CI_SUBMIT abap2xlsx Demo: Macro enabled workbook
  submit zdemo_excel30 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: ABAP Cell data types + leading blanks string
  submit zdemo_excel31 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Autosize Column with different Font sizes
  " zdemo_excel32 is not added because it uses ALV and cannot be processed (Native)
  submit zdemo_excel33 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Table autofilter
  submit zdemo_excel34 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Static Styles Chess
  submit zdemo_excel35 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Static Styles
  submit zdemo_excel36 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Style applied to sheet, column and single cell
  submit zdemo_excel37 with p_upfile  = lv_upfile
                       with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Simplest call of the reader and writer - passthrough data
  submit zdemo_excel38 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Show off integration of drawings ( here using the SAP-Icons )
  submit zdemo_excel39 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Charts
  submit zdemo_excel40 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Demo Printsettings
  submit zdemo_excel41 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Inheritance
  submit zdemo_excel44 with rb_down = abap_true with rb_show = abap_false with  p_path    = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: No line if empty

  submit zdemo_excel_comments     with rb_down = abap_true with rb_show = abap_false with  p_path = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Comments
  submit ztest_excel_image_header with rb_down = abap_true with rb_show = abap_false with  p_path = p_path and return. "#EC CI_SUBMIT abap2xlsx Demo: Image in Header and Footer
  "
  " Reader/Writer Demo must always run at the end
  " to make sure all documents where created
  "
  submit zdemo_excel15 with  p_path = p_path and return. "#EC CI_SUBMIT Read Excel and write it back
