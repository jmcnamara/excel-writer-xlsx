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
$worksheet->add_table( 'C3:F14', { total_row => 1 } );

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
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F14" totalsRowCount="1">
  <autoFilter ref="C3:F13"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Column1"/>
    <tableColumn id="2" name="Column2"/>
    <tableColumn id="3" name="Column3"/>
    <tableColumn id="4" name="Column4"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
