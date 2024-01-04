###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_worksheet);
use strict;
use warnings;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;


###############################################################################
#
# Test the _assemble_xml_file() method.
#
# Test conditional formats.
#
$caption = " \tWorksheet: _assemble_xml_file()";

$worksheet = _new_worksheet(\$got);

$worksheet->select();

# Start test code.
$worksheet->write( 'A1', 10 );
$worksheet->write( 'A2', 20 );
$worksheet->write( 'A3', 30 );
$worksheet->write( 'A4', 40 );

$worksheet->write( 'B1', 5 );

$worksheet->conditional_formatting( 'A1',
    {
        type     => 'cell',
        format   => undef,
        criteria => 'greater than',
        value    => '$B$1',
    }
);
# End test code.

$worksheet->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B4"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:2">
      <c r="A1">
        <v>10</v>
      </c>
      <c r="B1">
        <v>5</v>
      </c>
    </row>
    <row r="2" spans="1:2">
      <c r="A2">
        <v>20</v>
      </c>
    </row>
    <row r="3" spans="1:2">
      <c r="A3">
        <v>30</v>
      </c>
    </row>
    <row r="4" spans="1:2">
      <c r="A4">
        <v>40</v>
      </c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A1">
    <cfRule type="cellIs" priority="1" operator="greaterThan">
      <formula>$B$1</formula>
    </cfRule>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
