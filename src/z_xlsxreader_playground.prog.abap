REPORT Z_XLSXREADER_PLAYGROUND.

class lcl_app definition.

  public section.

    class-methods run
      raising
        cx_openxml_not_found
        cx_openxml_format
        cx_static_check .
endclass.

class lcl_app implementation.

  method run.

    data xdata type xstring.
    data xl type ref to zcl_xlsxreader.
    data sheets type string_table.
    data tab type zcl_xlsxreader=>tt_table.

    xdata  = zcl_w3mime_fs=>read_file_x( 'c:\tmp\Example.xlsx ' ).
    xl     = zcl_xlsxreader=>create( xdata ).
    sheets = xl->get_sheet_names( ).

    tab = xl->get_sheet( '_contents' ).
    tab = xl->get_sheet( 'TESTCASES' ).
    tab = xl->get_sheet( 'SFLIGHT' ).
    tab = xl->get_sheet( 'COMPLEX' ).

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
