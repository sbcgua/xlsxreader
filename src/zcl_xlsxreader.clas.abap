class ZCL_XLSXREADER definition
  public
  create public .

public section.

  types:
    begin of ty_raw_cell,
      index    type i,
      type     type c length 1,
      cell_ref type string,
      value    type string,
      column   type string,
      row      type string,
      style    type i,
    end of ty_raw_cell,
    tt_raw_cells type standard table of ty_raw_cell with default key.

  types:
    begin of ts_table,
      col   type c length 3,
      row   type i,
      type  type c length 1,
      value type string,
    end of ts_table .

  types:
    begin of ty_style,
      num_format type string,
    end of ty_style.

  types:
    tt_styles type standard table of ty_style with default key.

  types:
    tt_table type standard table of ts_table with key col row .

  class-methods load
    importing
      !iv_xdata type xstring
    returning
      value(ro_instance) type ref to zcl_xlsxreader
    raising
      cx_openxml_format
      cx_openxml_not_found.

  methods get_sheet
    importing
      !iv_name type string
    returning
      value(rt_table) type tt_table
    raising
      cx_openxml_not_found
      cx_openxml_format .

  methods constructor
    importing
      iv_xdata type xstring
    raising
      cx_openxml_format
      cx_openxml_not_found.

  methods get_sheet_names
    returning
      value(rt_sheet_names) type string_table
    raising
      cx_openxml_format .

  methods get_styles
    returning
      value(rt_styles) type tt_styles
    raising
      cx_openxml_format
      cx_openxml_not_found.

protected section.
private section.

  constants c_excldt type dats value '19000101' ##NO_TEXT.

  data mo_workbook type ref to cl_xlsx_workbookpart .
  data mt_sheets type zcl_xlsxreader_proc_sheets=>tt_sheets .
  data mo_xlsx type ref to cl_xlsx_document .
  data mt_shared_strings type string_table.

  methods get_sheets
    returning
      value(rt_sheets) type zcl_xlsxreader_proc_sheets=>tt_sheets
    raising
      cx_openxml_format .

  methods convert_date
    importing
      !iv_days type string
    returning
      value(rv_date) type dats .

  methods get_shared_string
    importing
      iv_index type i
    returning
      value(rv_str) type string.

  methods load_shared_strings
    raising
      cx_openxml_format cx_openxml_not_found.

  methods load_worksheet_raw
    importing
      iv_name type string
    returning
      value(rt_raw_cells) type tt_raw_cells
    raising
      cx_openxml_format cx_openxml_not_found.

  methods get_iterator_of
    importing
      io_xml_doc type ref to if_ixml_document
      iv_tag_name type string
    returning
      value(ro_iterator) type ref to if_ixml_node_iterator.

  methods parse_row
    importing
      io_node type ref to if_ixml_node
    changing
      ct_raw_cells type tt_raw_cells.

  methods add_default_num_formats
    changing ct_num_formats type zcl_xlsxreader_proc_num_fmts=>ts_num_formats.

ENDCLASS.



CLASS ZCL_XLSXREADER IMPLEMENTATION.


  method add_default_num_formats.

    data ls_num_format like line of ct_num_formats.

    define _add_num_format.
      ls_num_format-numfmtid   = &1.
      ls_num_format-formatcode = &2.
      insert ls_num_format into table ct_num_formats.
    end-of-definition.

    _add_num_format 0  'General'.
    _add_num_format 1  '0'.
    _add_num_format 2  '0.00'.
    _add_num_format 3  '#,##0'.
    _add_num_format 4  '#,##0.00'.
    _add_num_format 5  '$#,##0_);($#,##0)'.
    _add_num_format 6  '$#,##0_);[Red]($#,##0)'.
    _add_num_format 7  '$#,##0.00_);($#,##0.00)'.
    _add_num_format 8  '$#,##0.00_);[Red]($#,##0.00)'.
    _add_num_format 9  '0%'.
    _add_num_format 10 '0.00%'.
    _add_num_format 11 '0.00E+00'.
    _add_num_format 12 '# ?/?'.
    _add_num_format 13 '# ??/??'.
    _add_num_format 14 'm/d/yy'.
    _add_num_format 15 'd-mmm-yy'.
    _add_num_format 16 'd-mmm'.
    _add_num_format 17 'mmm-yy'.
    _add_num_format 18 'h:mm AM/PM'.
    _add_num_format 19 'h:mm:ss AM/PM'.
    _add_num_format 20 'h:mm'.
    _add_num_format 21 'h:mm:ss'.
    _add_num_format 22 'm/d/yy h:mm'.
    _add_num_format 36 'm/d/yy'.
    _add_num_format 37 '#,##0 ;(#,##0)'.
    _add_num_format 38 '#,##0 ;[Red](#,##0)'.
    _add_num_format 39 '#,##0.00;(#,##0.00)'.
    _add_num_format 40 '#,##0.00;[Red](#,##0.00)'.
    _add_num_format 45 'mm:ss'.
    _add_num_format 46 '[h]:mm:ss'.
    _add_num_format 47 'mmss.0'.
    _add_num_format 48 '##0.0E+0'.
    _add_num_format 49 '@'.
    _add_num_format 50 'm/d/yy'.
    _add_num_format 51 'm/d/yy'.
    _add_num_format 52 'm/d/yy'.
    _add_num_format 53 'm/d/yy'.
    _add_num_format 54 'm/d/yy'.
    _add_num_format 55 'm/d/yy'.
    _add_num_format 56 'm/d/yy'.
    _add_num_format 57 'm/d/yy'.
    _add_num_format 58 'm/d/yy'.
    _add_num_format 59 '0'.
    _add_num_format 60 '0.00'.
    _add_num_format 61 '#,##0'.
    _add_num_format 62 '#,##0.00'.
    _add_num_format 63 '$#,##0_);($#,##0)'.
    _add_num_format 64 '$#,##0_);[Red]($#,##0)'.
    _add_num_format 65 '$#,##0.00_);($#,##0.00)'.
    _add_num_format 66 '$#,##0.00_);[Red]($#,##0.00)'.
    _add_num_format 67 '0%'.
    _add_num_format 68 '0.00%'.
    _add_num_format 69 '# ?/?'.
    _add_num_format 70 '# ??/??'.
    _add_num_format 71 'm/d/yy'.
    _add_num_format 72 'm/d/yy'.
    _add_num_format 73 'd-mmm-yy'.
    _add_num_format 74 'd-mmm'.
    _add_num_format 75 'mmm-yy'.
    _add_num_format 76 'h:mm'.
    _add_num_format 77 'h:mm:ss'.
    _add_num_format 78 'm/d/yy h:mm'.
    _add_num_format 79 'mm:ss'.
    _add_num_format 80 '[h]:mm:ss'.
    _add_num_format 81 'mmss.0'.

  endmethod.


  method constructor.
    mo_xlsx = cl_xlsx_document=>load_document( iv_xdata ).
    mo_workbook = mo_xlsx->get_workbookpart( ).
  endmethod.


  method convert_date.
    data lv_days type i.

    check iv_days co '0123456789'.
    lv_days = iv_days.
    rv_date = c_excldt + lv_days.
  endmethod.


  method get_iterator_of.

    data lo_ixml_root type ref to if_ixml_element.
    data lo_nodes type ref to if_ixml_node_collection.

    lo_ixml_root = io_xml_doc->get_root_element( ).
    lo_nodes     = lo_ixml_root->get_elements_by_tag_name( name = iv_tag_name ).
    ro_iterator  = lo_nodes->create_iterator( ).

  endmethod.


  method get_shared_string.

    read table mt_shared_strings into rv_str index iv_index.

  endmethod.


  method get_sheet.

    data lt_raw_cells type tt_raw_cells.
    lt_raw_cells = load_worksheet_raw( iv_name ).
    load_shared_strings( ).

    field-symbols <c> like line of lt_raw_cells.
    field-symbols <res> like line of rt_table.

    " post process
    loop at lt_raw_cells assigning <c>.
      "get column
      <c>-column = <c>-cell_ref.
      condense <c>-row no-gaps.
      replace <c>-row in <c>-column with space.

      if <c>-type eq 's'.
        <c>-value = get_shared_string( <c>-index + 1 ).
      endif.
      condense <c>-value. " ???

      append initial line to rt_table assigning <res>.
      <res>-row   = <c>-row.
      <res>-col   = <c>-column.
      <res>-type  = <c>-type.
      <res>-value = <c>-value.
    endloop.

  endmethod.


  method get_sheets.

    if mt_sheets is initial.
      data lo_xml_doc type ref to if_ixml_document.
      lo_xml_doc = zcl_xlsxreader_xml_utils=>parse_xmldoc( mo_workbook->get_data( ) ).
      mt_sheets  = zcl_xlsxreader_proc_sheets=>read( lo_xml_doc ).
    endif.

    rt_sheets = mt_sheets.

  endmethod.


  method get_sheet_names.

    data lt_sheets like mt_sheets.
    field-symbols <s> like line of lt_sheets.

    lt_sheets = get_sheets( ).

    loop at lt_sheets assigning <s>.
      append <s>-name to rt_sheet_names.
    endloop.

  endmethod.


  method get_styles.

    data lo_style_part type ref to cl_xlsx_stylespart.
    data lo_xml_doc type ref to if_ixml_document.
    lo_style_part = mo_workbook->get_stylespart( ).
    lo_xml_doc    = zcl_xlsxreader_xml_utils=>parse_xmldoc( lo_style_part->get_data( ) ).

    data lt_num_formats type zcl_xlsxreader_proc_num_fmts=>ts_num_formats.
    field-symbols <num_format> like line of lt_num_formats.
    lt_num_formats = zcl_xlsxreader_proc_num_fmts=>read( lo_xml_doc ).
    add_default_num_formats( changing ct_num_formats = lt_num_formats ).

    data lt_cell_styles type zcl_xlsxreader_proc_styles=>tt_cell_styles.
    field-symbols <cell_style> like line of lt_cell_styles.
    lt_cell_styles = zcl_xlsxreader_proc_styles=>read( lo_xml_doc ).

    field-symbols <style> like line of rt_styles.
    append initial line to rt_styles. " Default standard style
    append initial line to rt_styles. " Default standard style 2 ???

    loop at lt_cell_styles assigning <cell_style>.
      append initial line to rt_styles assigning <style>.
      read table lt_num_formats assigning <num_format> with key numfmtid = <cell_style>-numfmtid.
      if sy-subrc = 0.
        <style>-num_format = <num_format>-formatcode.
      endif.
    endloop.

  endmethod.


  method load.

    create object ro_instance
      exporting
        iv_xdata = iv_xdata.

  endmethod.


  method load_shared_strings.

    if lines( mt_shared_strings ) > 0.
      return.
    endif.

    data lo_shared_st type ref to cl_xlsx_sharedstringspart.
    data lo_xml_doc type ref to if_ixml_document.

    lo_shared_st = mo_workbook->get_sharedstringspart( ).
    lo_xml_doc   = zcl_xlsxreader_xml_utils=>parse_xmldoc( lo_shared_st->get_data( ) ).
    mt_shared_strings = zcl_xlsxreader_proc_shared_str=>read( lo_xml_doc ).

  endmethod.


  method load_worksheet_raw.

    data ls_sheet like line of mt_sheets.
    data lo_worksheet type ref to cl_xlsx_worksheetpart.
    data lo_node_iterator type ref to if_ixml_node_iterator.
    data lo_node type ref to if_ixml_node.

    read table mt_sheets into ls_sheet with table key name = iv_name.
    if sy-subrc ne 0.
      raise exception type cx_openxml_not_found.
    endif.
    lo_worksheet    ?= mo_workbook->get_part_by_id( ls_sheet-id ).
    lo_node_iterator = get_iterator_of(
      io_xml_doc  = zcl_xlsxreader_xml_utils=>parse_xmldoc( lo_worksheet->get_data( ) )
      iv_tag_name = 'row' ).
    lo_node          = lo_node_iterator->get_next( ).

    while lo_node is not initial.
      parse_row(
        exporting
          io_node = lo_node
        changing
          ct_raw_cells = rt_raw_cells ).
      lo_node          = lo_node_iterator->get_next( ).
    endwhile.

  endmethod.


  method parse_row.

    field-symbols <c> like line of ct_raw_cells.
    data lo_attrs type ref to if_ixml_named_node_map.
    data lo_node_iterator type ref to if_ixml_node_iterator.
    data lo_cell_node type ref to if_ixml_node.
    data lo_attr type ref to if_ixml_node.
    data lv_row like <c>-row.

    lo_attrs         = io_node->get_attributes( ).
    lv_row           = lo_attrs->get_named_item( 'r' )->get_value( ).
    lo_node_iterator = io_node->get_children( )->create_iterator( ).
    lo_cell_node     = lo_node_iterator->get_next( ).

    while lo_cell_node is bound.
      append initial line to ct_raw_cells assigning <c>.
      <c>-row = lv_row.

      lo_attrs     = lo_cell_node->get_attributes( ).
      <c>-cell_ref = lo_attrs->get_named_item( 'r' )->get_value( ).

      lo_attr = lo_attrs->get_named_item( 't' ).
      if lo_attr is bound.
        <c>-type = lo_attr->get_value( ).
      endif.

      lo_attr = lo_attrs->get_named_item( 's' ).
      if lo_attr is bound.
        <c>-style = lo_attr->get_value( ).
      endif.

      if <c>-type = 's'. " string
        <c>-index = lo_cell_node->get_value( ).
      else.
        <c>-value = lo_cell_node->get_value( ).
      endif.

      lo_cell_node     = lo_node_iterator->get_next( ).
    endwhile.

  endmethod.
ENDCLASS.
