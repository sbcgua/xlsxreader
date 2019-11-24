REPORT Z_XLSXREADER_PERFORMANCE.

include ztest_benchmark.

class lcl_app definition.

  public section.

    class-methods main
      raising
        cx_openxml_not_found
        cx_openxml_format
        cx_static_check .

    methods prepare.
    methods xlsx_reader.
    methods abap2xlsx.
    methods run
      importing
        iv_method type string.

    data mv_xdata type xstring.
    data mv_num_rounds type i.

endclass.

class lcl_app implementation.

  method prepare.
    mv_xdata  = zcl_w3mime_fs=>read_file_x( 'c:\tmp\Example.xlsx ' ).
  endmethod.

  method xlsx_reader.

    data xl type ref to zcl_xlsxreader.
    data sheets type string_table.
    data tab type zcl_xlsxreader=>tt_cells.
    data styles type zcl_xlsxreader=>tt_styles.

    xl     = zcl_xlsxreader=>load( mv_xdata ).
    styles = xl->get_styles( ).
    sheets = xl->get_sheet_names( ).
    tab = xl->get_sheet( '_contents' ).
    tab = xl->get_sheet( 'TESTCASES' ).
    tab = xl->get_sheet( 'SFLIGHT' ).
    tab = xl->get_sheet( 'COMPLEX' ).

  endmethod.

  method abap2xlsx.

    data lo_reader type ref to zif_excel_reader.
    data lo_excel type ref to zcl_excel.

    create object lo_reader type zcl_excel_reader_2007.
    lo_excel = lo_reader->load( mv_xdata ).

    " styles
    data:
      lv_tmp   type string,
      lo_style type ref to zcl_excel_style,
      lo_iter  type ref to cl_object_collection_iterator.

    lo_iter = lo_excel->get_styles_iterator( ).
    while lo_iter->has_next( ) is not initial.
      lo_style ?= lo_iter->get_next( ).
      lv_tmp = lo_style->get_guid( ).
      lv_tmp = lo_style->number_format->format_code.
    endwhile.

    " sheet names + content
    data lo_worksheet type ref to zcl_excel_worksheet.
    lo_iter = lo_excel->get_worksheets_iterator( ).
    while lo_iter->has_next( ) is not initial.
      lo_worksheet ?= lo_iter->get_next( ).
      lv_tmp        = lo_worksheet->get_title( ).
*      worksheet->sheet_content
    endwhile.

  endmethod.

  method run.

    data lo_benchmark type ref to lcl_benchmark.

    create object lo_benchmark
      exporting
        io_object = me
        iv_method = iv_method
        iv_times  = mv_num_rounds.

    lo_benchmark->run( ).
    lo_benchmark->print( ).

  endmethod.

  method main.

    data lo_app type ref to lcl_app.
    create object lo_app.

    lo_app->mv_num_rounds = 100.
    lo_app->prepare( ).
    lo_app->run( 'xlsx_reader' ).
    lo_app->run( 'abap2xlsx' ).

  endmethod.

endclass.


start-of-selection.

  data gx type ref to cx_root.
  try .
    lcl_app=>main( ).
  catch cx_root into gx.
    data gmsg type string.
    gmsg = gx->get_text( ).
    write: / gmsg.
  endtry.
