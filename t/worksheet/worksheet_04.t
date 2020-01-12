###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
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
# Test row formatting.
#
$caption = " \tWorksheet: _assemble_xml_file()";

my $format = Excel::Writer::XLSX::Format->new( {}, {}, xf_index => 1, bold => 1 );
$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_row( 1, 30 );
$worksheet->set_row( 3, undef, undef, 1 );
$worksheet->set_row( 6, undef, $format );
$worksheet->set_row( 9, 3 );
$worksheet->set_row( 12, 24, undef, 1 );
$worksheet->set_row( 14, 0 );


$worksheet->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A2:A15"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="2" ht="30" customHeight="1"/>
    <row r="4" hidden="1"/>
    <row r="7" s="1" customFormat="1"/>
    <row r="10" ht="3" customHeight="1"/>
    <row r="13" ht="24" hidden="1" customHeight="1"/>
    <row r="15" hidden="1"/>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
