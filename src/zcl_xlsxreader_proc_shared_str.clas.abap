class ZCL_XLSXREADER_PROC_SHARED_STR definition
  public
  final
  create private .


  public section.
    interfaces zif_xlsxreader_node_processor.
    class-methods read
      importing
        io_xml_doc type ref to if_ixml_document
      returning
        value(rt_shared_strings) type string_table.

  protected section.
  private section.
    data mt_shared_strings type string_table.
ENDCLASS.



CLASS ZCL_XLSXREADER_PROC_SHARED_STR IMPLEMENTATION.


  method read.

    data lo_processor type ref to zcl_xlsxreader_proc_shared_str.
    create object lo_processor.
    zcl_xlsxreader_xml_utils=>iterate_children(
      io_node = io_xml_doc->find_from_name_ns(
        name = 'sst'
        uri = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' )
      ii_item_processor = lo_processor ).
    rt_shared_strings = lo_processor->mt_shared_strings.

  endmethod.


  method zif_xlsxreader_node_processor~process_node.

    data lv_str type string.
    lv_str = io_node->get_value( ).
    append lv_str to mt_shared_strings.

  endmethod.
ENDCLASS.
