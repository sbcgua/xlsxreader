class ZCL_XLSXREADER_PROC_STYLES definition
  public
  final
  create private .

  public section.

    interfaces zif_xlsxreader_node_processor .

    types:
      begin of ty_cell_style,
        numfmtid   type i,
      end of ty_cell_style .
    types:
      tt_cell_styles type table of ty_cell_style with default key .

    class-methods read
      importing
        !io_xml_doc type ref to if_ixml_document
      returning
        value(rt_cell_styles) type tt_cell_styles .
  protected section.
  private section.
    data mt_cell_styles type tt_cell_styles.
ENDCLASS.



CLASS ZCL_XLSXREADER_PROC_STYLES IMPLEMENTATION.


  method READ.

    data lo_processor type ref to zcl_xlsxreader_proc_styles.
    create object lo_processor.

    zcl_xlsxreader_xml_utils=>iterate_children(
      io_node = io_xml_doc->find_from_name_ns(
        name = 'cellXfs'
        uri  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' )
      ii_item_processor = lo_processor ).
    rt_cell_styles = lo_processor->mt_cell_styles.

  endmethod.


  method ZIF_XLSXREADER_NODE_PROCESSOR~PROCESS_NODE.

    data ls_cell_style like line of mt_cell_styles.

    zcl_xlsxreader_xml_utils=>attributes_to_struc(
      exporting
        io_node = io_node
      importing
        es_struc = ls_cell_style ).

    append ls_cell_style to mt_cell_styles.

  endmethod.
ENDCLASS.
