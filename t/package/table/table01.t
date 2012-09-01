###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Table methods.
#
# reverse('(c)'), September 2012, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::Table;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Package::Table' );

my %properties = (
    _autofilter => 'C3:F13',
    _columns    => [
        [ 1, 'Column1' ],
        [ 2, 'Column2' ],
        [ 3, 'Column3' ],
        [ 4, 'Column4' ],
    ],
    _style            => 'TableStyleMedium9',
    _show_first_col   => 0,
    _show_last_col    => 0,
    _show_row_stripes => 1,
    _show_col_stripes => 0,
);


###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tTable: _assemble_xml_file()";

$obj->_set_properties( \%properties );
$obj->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
  <autoFilter ref="C3:F13"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Column1"/>
    <tableColumn id="2" name="Column2"/>
    <tableColumn id="3" name="Column3"/>
    <tableColumn id="4" name="Column4"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
