interface zif_xlsxreader_node_processor
  public.

  methods process_node
    importing
      io_node type ref to if_ixml_node
      i_context type any optional.

endinterface.
