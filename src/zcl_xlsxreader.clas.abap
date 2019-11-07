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
      style    type string,
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
    begin of ts_sheet,
      name  type string,
      id    type string,
    end of ts_sheet .

  types:
    tt_table type standard table of ts_table with key col row .
  types:
    tt_sheet type standard table of ts_sheet with key name .

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

protected section.
private section.

  data m_workbook type ref to cl_xlsx_workbookpart .
  data m_sheets type tt_sheet .
  data m_xlsx type ref to cl_xlsx_document .
  constants c_ns_r type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' ##NO_TEXT.
  constants c_excldt type dats value '19000101' ##NO_TEXT.
  data mt_shared_strings type string_table.

  methods get_sheets
    returning
      value(rt_sheets) type tt_sheet
    raising
      cx_openxml_format .

  methods get_xmldoc
    importing
      !iv_xml type xstring
    returning
      value(ro_xmldoc) type ref to if_ixml_document .

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

ENDCLASS.



CLASS ZCL_XLSXREADER IMPLEMENTATION.


  method constructor.
    m_xlsx = cl_xlsx_document=>load_document( iv_xdata ).
    m_workbook = m_xlsx->get_workbookpart( ).
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

    data ls_sheet type ts_sheet.
    data lo_node_iterator type ref to if_ixml_node_iterator.
    data lo_node type ref to if_ixml_node.
    data lo_attrs type ref to if_ixml_named_node_map.

    if m_sheets is initial.
      lo_node_iterator = get_iterator_of(
        io_xml_doc  = get_xmldoc( m_workbook->get_data( ) )
        iv_tag_name = 'sheet' ).
      lo_node = lo_node_iterator->get_next( ).

      while lo_node is bound.
        lo_attrs      = lo_node->get_attributes( ).
        ls_sheet-name = lo_attrs->get_named_item( 'name' )->get_value( ).
        ls_sheet-id   = lo_attrs->get_named_item_ns(
          name = 'id'
          uri  = c_ns_r )->get_value( ).
        append ls_sheet to me->m_sheets.
        lo_node = lo_node_iterator->get_next( ).
      endwhile.
    endif.

    rt_sheets = m_sheets.

  endmethod.


  method get_sheet_names.

    data lt_sheets like m_sheets.
    field-symbols <s> like line of lt_sheets.

    lt_sheets = get_sheets( ).

    loop at lt_sheets assigning <s>.
      append <s>-name to rt_sheet_names.
    endloop.

  endmethod.


  method get_xmldoc.

    data lo_ixml type ref to if_ixml.
    data lo_ixml_sf type ref to if_ixml_stream_factory.
    data lo_ixml_stream type ref to if_ixml_istream.
    data lo_ixml_parser type ref to if_ixml_parser.

    lo_ixml        = cl_ixml=>create( ).
    lo_ixml_sf     = lo_ixml->create_stream_factory( ).
    lo_ixml_stream = lo_ixml_sf->create_istream_xstring( iv_xml ).
    ro_xmldoc      = lo_ixml->create_document( ).
    lo_ixml_parser = lo_ixml->create_parser(
      document       = ro_xmldoc
      istream        = lo_ixml_stream
      stream_factory = lo_ixml_sf ).

    lo_ixml_parser->parse( ).

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
    data lo_node_iterator type ref to if_ixml_node_iterator.
    data lo_node type ref to if_ixml_node.
    data lv_str type string.

    lo_shared_st     = m_workbook->get_sharedstringspart( ).
    lo_node_iterator = get_iterator_of(
      io_xml_doc  = get_xmldoc( lo_shared_st->get_data( ) )
      iv_tag_name = 'si' ).
    lo_node = lo_node_iterator->get_next( ).
    while lo_node is not initial.
      lv_str = lo_node->get_value( ).
      append lv_str to mt_shared_strings.
      lo_node = lo_node_iterator->get_next( ).
    endwhile.

  endmethod.


  method load_worksheet_raw.

    data ls_sheet type ts_sheet.
    data lo_worksheet type ref to cl_xlsx_worksheetpart.
    data lo_node_iterator type ref to if_ixml_node_iterator.
    data lo_node type ref to if_ixml_node.

    read table m_sheets into ls_sheet with table key name = iv_name.
    if sy-subrc ne 0.
      raise exception type cx_openxml_not_found.
    endif.
    lo_worksheet    ?= m_workbook->get_part_by_id( ls_sheet-id ).
    lo_node_iterator = get_iterator_of(
      io_xml_doc  = get_xmldoc( lo_worksheet->get_data( ) )
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
