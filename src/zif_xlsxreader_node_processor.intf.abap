interface ZIF_XLSXREADER_NODE_PROCESSOR
  public .

  methods process_node
    importing
      io_node type ref to if_ixml_node.

endinterface.
