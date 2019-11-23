class ZCL_XLSXREADER_PROC_SHEETS definition
  public
  final
  create private .

  public section.

    interfaces zif_xlsxreader_node_processor .

    types:
      begin of ty_sheet,
        name    type string,
        sheetid type i,
        id      type string,
      end of ty_sheet .
    types:
      tt_sheets type table of ty_sheet with key name .

    class-methods read
      importing
        !io_xml_doc type ref to if_ixml_document
      returning
        value(rt_sheets) type tt_sheets .
  protected section.
  private section.
    data mt_sheets type tt_sheets.
ENDCLASS.



CLASS ZCL_XLSXREADER_PROC_SHEETS IMPLEMENTATION.


  method READ.

    data lo_processor type ref to zcl_xlsxreader_proc_sheets.
    create object lo_processor.

    zcl_xlsxreader_xml_utils=>iterate_children(
      io_node = io_xml_doc->find_from_name_ns(
        name = 'sheets'
        uri  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' )
      ii_item_processor = lo_processor ).
    rt_sheets = lo_processor->mt_sheets.

  endmethod.


  method ZIF_XLSXREADER_NODE_PROCESSOR~PROCESS_NODE.

    data ls_sheet like line of mt_sheets.

    zcl_xlsxreader_xml_utils=>attributes_to_struc(
      exporting
        io_node = io_node
      importing
        es_struc = ls_sheet ).

    append ls_sheet to mt_sheets.

  endmethod.
ENDCLASS.
