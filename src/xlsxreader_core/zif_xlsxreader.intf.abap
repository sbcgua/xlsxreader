interface zif_xlsxreader
  public.

  constants origin type string value 'https://github.com/sbcgua/xlsxreader'. "#EC NOTEXT
  constants origin_forked_from type string value 'https://github.com/mkysoft/xlsxreader'. "#EC NOTEXT
  constants license type string value 'MIT'. "#EC NOTEXT

  constants c_openxml_namespace_uri type string
    value 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' ##NO_TEXT.

  types:
    begin of ty_num_format,
      numfmtid   type i,
      formatcode type string,
    end of ty_num_format.
  types:
    tt_num_formats type table of ty_num_format with key numfmtid.
  types:
    ts_num_formats type sorted table of ty_num_format with unique key numfmtid.
  types:
    begin of ty_sheet,
      name    type string,
      sheetid type i,
      id      type string,
    end of ty_sheet.
  types:
    tt_sheets type table of ty_sheet with key name.
  types:
    begin of ty_cell_style,
      numfmtid   type i,
    end of ty_cell_style.
  types:
    tt_cell_styles type table of ty_cell_style with default key.
  types:
    begin of ty_raw_cell,
      r     type string,
      s     type i,
      t     type string,
      row   type string,
      value type string,
    end of ty_raw_cell.
  types:
    tt_raw_cells type standard table of ty_raw_cell with key r.
  types:
    begin of ty_parsing_context,
      stage type string,
      data  type ref to data,
    end of ty_parsing_context.
  types:
    begin of ty_cell,
      col   type i,
      row   type i,
      type  type c length 1,
      style type i,
      value type string,
      ref   type string,
    end of ty_cell.
  types:
    begin of ty_style,
      num_format type string,
    end of ty_style.
  types:
    tt_styles type standard table of ty_style with default key.
  types:
    tt_cells type standard table of ty_cell with key col row.

**********************************************************************
* METHODS
**********************************************************************

  methods get_sheet
    importing
      !iv_name type string
    returning
      value(rt_cells) type tt_cells
    raising
      cx_openxml_not_found
      cx_openxml_format.

  methods get_sheet_names
    returning
      value(rt_sheet_names) type string_table
    raising
      cx_openxml_format.

  methods get_styles
    returning
      value(rt_styles) type tt_styles
    raising
      cx_openxml_format
      cx_openxml_not_found.

endinterface.
