class ZCL_XLSXREADER_XML_UTILS definition
  public
  final
  create public .

  public section.

    class-methods iterate_children
      importing
        io_node type ref to if_ixml_node
        ii_item_processor type ref to zif_xlsxreader_node_processor.

    class-methods iterate_children_by_tag_name
      importing
        io_element type ref to if_ixml_element
        iv_tag_name type string
        ii_item_processor type ref to zif_xlsxreader_node_processor.

    class-methods parse_xmldoc
      importing
        !iv_xml type xstring
      returning
        value(ro_xmldoc) type ref to if_ixml_document .

    class-methods attributes_to_struc
      importing
        io_node type ref to if_ixml_node
      exporting
        es_struc type any.

  protected section.
  private section.
    class-methods iterate_nodes
      importing
        io_node_iterator type ref to if_ixml_node_iterator
        ii_item_processor type ref to zif_xlsxreader_node_processor.

ENDCLASS.



CLASS ZCL_XLSXREADER_XML_UTILS IMPLEMENTATION.


  method attributes_to_struc.

    data lo_attrs     type ref to if_ixml_named_node_map.
    data lo_attr      type ref to if_ixml_attribute.
    data lo_iterator  type ref to if_ixml_node_iterator.
    data lv_attr_name type string.
    field-symbols <fld> type any.

    clear es_struc.

    lo_attrs    = io_node->get_attributes( ).
    lo_iterator = lo_attrs->create_iterator( ).
    lo_attr    ?= lo_iterator->get_next( ).
    while lo_attr is bound.
      lv_attr_name = to_upper( lo_attr->get_name( ) ).
      assign component lv_attr_name of structure es_struc to <fld>.
      if sy-subrc = 0.
        <fld> = lo_attr->get_value( ).
      endif.
      lo_attr ?= lo_iterator->get_next( ).
    endwhile.

  endmethod.


  method iterate_children.
    iterate_nodes(
      io_node_iterator = io_node->get_children( )->create_iterator( )
      ii_item_processor = ii_item_processor ).
  endmethod.


  method iterate_children_by_tag_name.
    iterate_nodes(
      io_node_iterator = io_element->get_elements_by_tag_name_ns( name = iv_tag_name )->create_iterator( )
      ii_item_processor = ii_item_processor ).
  endmethod.


  method iterate_nodes.
    data lo_node type ref to if_ixml_node.
    lo_node          = io_node_iterator->get_next( ).

    while lo_node is bound.
      ii_item_processor->process_node( lo_node ).
      lo_node = io_node_iterator->get_next( ).
    endwhile.
  endmethod.


  method parse_xmldoc.

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
ENDCLASS.
