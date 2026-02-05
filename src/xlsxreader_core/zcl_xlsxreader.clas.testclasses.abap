class ltcl_xlreader definition final
  for testing
  duration short
  risk level harmless.
  private section.

    methods column_to_index for testing.
    methods read_test_data returning value(rv_xdata) type xstring.
    methods load for testing raising cx_static_check.
endclass.

class zcl_xlsxreader definition local friends ltcl_xlreader.

class ltcl_xlreader implementation.

  method column_to_index.

    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'A' )
      exp = 1 ).
    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'B' )
      exp = 2 ).
    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'Z' )
      exp = 26 ).
    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'AA' )
      exp = 27 ).
    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'AZ' )
      exp = 52 ).

  endmethod.

  method read_test_data.

    data lt_data type lvc_t_mime.
    data lv_size type i.

    data lv_filesize type w3_qvalue.
    data ls_object type wwwdatatab.

    ls_object-relid = 'MI'.
    ls_object-objid = 'ZXLSXREADER_UNIT_TEST'.

    call function 'WWWPARAMS_READ'
      exporting
        relid = ls_object-relid
        objid = ls_object-objid
        name  = 'filesize'
      importing
        value = lv_filesize
      exceptions
        others = 1.

    if sy-subrc > 0.
      cl_abap_unit_assert=>fail( 'Cannot read W3MI filesize' ).
    endif.

    lv_size = lv_filesize.

    call function 'WWWDATA_IMPORT'
      exporting
        key               = ls_object
      tables
        mime              = lt_data
      exceptions
        wrong_object_type = 1
        import_error      = 2.

    if sy-subrc > 0.
      cl_abap_unit_assert=>fail( 'Cannot read W3MI data' ).
    endif.

    call function 'SCMS_BINARY_TO_XSTRING'
      exporting
        input_length = lv_size
      importing
        buffer       = rv_xdata
      tables
        binary_tab   = lt_data.

  endmethod.

  method load.

    data lo_excel type ref to zcl_xlsxreader.
    lo_excel = zcl_xlsxreader=>load( read_test_data( ) ).

    " Sheet names
    data lt_sheets_exp type string_table.
    append 'Sheet1' to lt_sheets_exp.
    append 'Sheet2' to lt_sheets_exp.
    cl_abap_unit_assert=>assert_equals(
      act = lo_excel->get_sheet_names( )
      exp = lt_sheets_exp ).

    " Styles
    data lt_styles_act type zif_xlsxreader=>tt_styles.
    data ls_style like line of lt_styles_act.

    lt_styles_act = lo_excel->get_styles( ).
    read table lt_styles_act into ls_style index 6.

    cl_abap_unit_assert=>assert_equals(
      act = lines( lt_styles_act )
      exp = 6 ).
    cl_abap_unit_assert=>assert_equals(
      act = ls_style-num_format
      exp = 'm/d/yy' ). " mm-dd-yy ?

    " Content
    data lt_content_act type zif_xlsxreader=>tt_cells.
    data lt_content_exp type zif_xlsxreader=>tt_cells.
    field-symbols <i> like line of lt_content_exp.

    define _add_cell.
      append initial line to lt_content_exp assigning <i>.
      <i>-row     = &1.
      <i>-col     = &2.
      <i>-value   = &3.
      <i>-ref     = &4.
      <i>-style   = &5.
      <i>-type    = &6.
    end-of-definition.

    clear lt_content_exp.
    _add_cell 1 1 'Column1'   'A1' 4 's'.
    _add_cell 1 2 'Column2'   'B1' 4 's'.
    _add_cell 2 1 'A'         'A2' 0 's'.
    _add_cell 2 2 '1'         'B2' 5 ''.
    _add_cell 3 1 'B'         'A3' 0 's'.
    _add_cell 3 2 '2'         'B3' 0 ''.
    _add_cell 4 1 'C'         'A4' 0 's'.
    _add_cell 4 2 '3'         'B4' 0 ''.
    _add_cell 6 1 'More'      'A6' 0 's'.
    _add_cell 6 2 'Data'      'B6' 0 's'.

    lt_content_act = lo_excel->get_sheet( 'Sheet1' ).
    cl_abap_unit_assert=>assert_equals( act = lt_content_act exp = lt_content_exp ).

    clear lt_content_exp.
    _add_cell 1 1 'A'     'A1' 4 's'.
    _add_cell 1 2 'B'     'B1' 4 's'.
    _add_cell 1 3 'C'     'C1' 4 's'.
    _add_cell 1 4 'D'     'D1' 4 's'.
    _add_cell 2 1 'Vasya' 'A2' 0 's'.
    _add_cell 2 2 '43344' 'B2' 6 ''.
    _add_cell 2 3 '15'    'C2' 0 ''.
    _add_cell 2 4 '1'     'D2' 0 'b'.
    _add_cell 3 1 'Petya' 'A3' 0 's'.
    _add_cell 3 2 '43345' 'B3' 6 ''.
    _add_cell 3 3 '16.37' 'C3' 0 ''.
    _add_cell 3 4 '0'     'D3' 0 'b'.

    lt_content_act = lo_excel->get_sheet( 'Sheet2' ).
    cl_abap_unit_assert=>assert_equals( act = lt_content_act exp = lt_content_exp ).

  endmethod.

endclass.
