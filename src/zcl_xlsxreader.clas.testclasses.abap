class ltcl_xlreader definition final
  for testing
  duration short
  risk level harmless.
  private section.

    methods column_to_index for testing.

endclass.

class zcl_xlsxreader definition local friends ltcl_xlreader.

class ltcl_xlreader implementation.

  method column_to_index.

    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'A' )
      exp = 1 ).
    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'B' )
      exp = 2 ).
    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'Z' )
      exp = 26 ).
    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'AA' )
      exp = 27 ).
    cl_abap_unit_assert=>assert_equals(
      act = zcl_xlsxreader=>column_to_index( 'AZ' )
      exp = 52 ).

  endmethod.

endclass.
