class ZCL_XLSXREADER_PROC_NUM_FMTS definition
  public
  final
  create private .

  public section.

    types:
      begin of ty_num_format,
        numfmtid   type i,
        formatcode type string,
      end of ty_num_format.

    types:
      tt_num_formats type table of ty_num_format with key numfmtid.

    types:
      ts_num_formats type sorted table of ty_num_format with unique key numfmtid.

    interfaces zif_xlsxreader_node_processor .

    class-methods read
      importing
        !io_xml_doc type ref to if_ixml_document
      returning
        value(rt_num_formats) type tt_num_formats .
  protected section.
  private section.
    data mt_num_formats type tt_num_formats.
ENDCLASS.



CLASS ZCL_XLSXREADER_PROC_NUM_FMTS IMPLEMENTATION.


  method READ.

    data lo_processor type ref to zcl_xlsxreader_proc_num_fmts.
    create object lo_processor.

    zcl_xlsxreader_xml_utils=>iterate_children(
      io_node = io_xml_doc->find_from_name_ns(
        name = 'numFmts'
        uri  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' )
      ii_item_processor = lo_processor ).
    rt_num_formats = lo_processor->mt_num_formats.

  endmethod.


  method ZIF_XLSXREADER_NODE_PROCESSOR~PROCESS_NODE.

    data ls_num_format like line of mt_num_formats.

    zcl_xlsxreader_xml_utils=>attributes_to_struc(
      exporting
        io_node = io_node
      importing
        es_struc = ls_num_format ).

    append ls_num_format to mt_num_formats.

  endmethod.
ENDCLASS.
