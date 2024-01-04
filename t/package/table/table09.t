###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Table methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::Table;
use Excel::Writer::XLSX::Worksheet;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Package::Table' );
my $worksheet = Excel::Writer::XLSX::Worksheet->new();

# Set the table properties.
$worksheet->add_table(
    'B2:K8',
    {
        total_row => 1,
        columns => [
            { total_string => 'Total' },
            {},
            { total_function => 'average' },
            { total_function => 'count' },
            { total_function => 'count_nums' },
            { total_function => 'max' },
            { total_function => 'min' },
            { total_function => 'sum' },
            { total_function => 'std_dev' },
            { total_function => 'var' }
          ],

    }
);

$worksheet->_prepare_tables( 1, {} );


###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tTable: _assemble_xml_file()";

$obj->_set_properties( $worksheet->{_tables}->[0] );
$obj->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="B2:K8" totalsRowCount="1">
  <autoFilter ref="B2:K7"/>
  <tableColumns count="10">
    <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
    <tableColumn id="2" name="Column2"/>
    <tableColumn id="3" name="Column3" totalsRowFunction="average"/>
    <tableColumn id="4" name="Column4" totalsRowFunction="count"/>
    <tableColumn id="5" name="Column5" totalsRowFunction="countNums"/>
    <tableColumn id="6" name="Column6" totalsRowFunction="max"/>
    <tableColumn id="7" name="Column7" totalsRowFunction="min"/>
    <tableColumn id="8" name="Column8" totalsRowFunction="sum"/>
    <tableColumn id="9" name="Column9" totalsRowFunction="stdDev"/>
    <tableColumn id="10" name="Column10" totalsRowFunction="var"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
