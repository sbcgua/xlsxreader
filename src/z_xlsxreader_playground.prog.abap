REPORT Z_XLSXREADER_PLAYGROUND.

class lcl_app definition.

  public section.

    class-methods run
      raising
        cx_openxml_not_found
        cx_openxml_format
        cx_static_check .

    class-methods print_sheet
      importing
        tab type zcl_xlsxreader=>tt_cells.
endclass.

class lcl_app implementation.

  method run.

    data xdata type xstring.
    data xl type ref to zcl_xlsxreader.
    data sheets type string_table.
    data tab type zcl_xlsxreader=>tt_cells.
    data str type string.

    xdata  = zcl_w3mime_fs=>read_file_x( 'c:\tmp\Example.xlsx ' ).
    xl     = zcl_xlsxreader=>load( xdata ).
    sheets = xl->get_sheet_names( ).
    loop at sheets into str.
      write: / str.
    endloop.

*    data lo_reader type ref to zif_excel_reader.
*    data lo_excel type ref to zcl_excel.
*    create object lo_reader type zcl_excel_reader_2007.
*    lo_excel = lo_reader->load( xdata ).

    data styles type zcl_xlsxreader=>tt_styles.
    styles = xl->get_styles( ).

    tab = xl->get_sheet( '_contents' ).
    write: / '_contents'.
    print_sheet( tab ).

    tab = xl->get_sheet( 'TESTCASES' ).
    tab = xl->get_sheet( 'SFLIGHT' ).
    tab = xl->get_sheet( 'COMPLEX' ).
    write: / 'COMPLEX'.
    print_sheet( tab ).



  endmethod.

  method print_sheet.

    data cell like line of tab.
    data lo_struc type ref to cl_abap_structdescr.
    lo_struc ?= cl_abap_structdescr=>describe_by_data( cell ).

    field-symbols <field> like line of lo_struc->components.
    field-symbols <val> type any.
    new-line.
    loop at lo_struc->components assigning <field>.
      write (10) <field>-name.
    endloop.

    loop at tab into cell.
      new-line.
      loop at lo_struc->components assigning <field>.
        assign component <field>-name of structure cell to <val>.
        write (10) <val> left-justified.
      endloop.
    endloop.

  endmethod.

endclass.


start-of-selection.

  data gx type ref to cx_root.
  try .
    lcl_app=>run( ).
  catch cx_root into gx.
    data gmsg type string.
    gmsg = gx->get_text( ).
    write: / gmsg.
  endtry.
