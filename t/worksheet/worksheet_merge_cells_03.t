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
# Test merged cells.
#
$caption = " \tWorksheet: _assemble_xml_file()";

my $format1 = Excel::Writer::XLSX::Format->new( {}, {}, xf_index => 1 );
my $format2 = Excel::Writer::XLSX::Format->new( {}, {}, xf_index => 2 );
my $format3 = Excel::Writer::XLSX::Format->new( {}, {}, xf_index => 3 );
$worksheet = _new_worksheet(\$got);

$worksheet->set_column('B:C', 12);
$worksheet->{_date_1904} = 0;

$worksheet->select();
$worksheet->merge_range_type( 'formula',     'B14:C14', '=1+2',                 $format1, 3 );
$worksheet->merge_range_type( 'number',      'B2:C2',   123,                    $format1 );
$worksheet->merge_range_type( 'string',      'B4:C4',   'foo',                  $format1 );
$worksheet->merge_range_type( 'blank',       'B6:C6',                           $format1 );
$worksheet->merge_range_type( 'rich_string', 'B8:C8',   'This is ', $format2, 'bold', $format1 );
$worksheet->merge_range_type( 'date_time',   'B10:C10', '2011-01-01T',          $format2 );
$worksheet->merge_range_type( 'url',         'B12:C12', 'http://www.perl.com/', $format3 );


$worksheet->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="B2:C14"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="2" max="3" width="12.7109375" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="2" spans="2:3">
      <c r="B2" s="1">
        <v>123</v>
      </c>
      <c r="C2" s="1"/>
    </row>
    <row r="4" spans="2:3">
      <c r="B4" s="1" t="s">
        <v>0</v>
      </c>
      <c r="C4" s="1"/>
    </row>
    <row r="6" spans="2:3">
      <c r="B6" s="1"/>
      <c r="C6" s="1"/>
    </row>
    <row r="8" spans="2:3">
      <c r="B8" s="1" t="s">
        <v>1</v>
      </c>
      <c r="C8" s="1"/>
    </row>
    <row r="10" spans="2:3">
      <c r="B10" s="2">
        <v>40544</v>
      </c>
      <c r="C10" s="2"/>
    </row>
    <row r="12" spans="2:3">
      <c r="B12" s="3" t="s">
        <v>2</v>
      </c>
      <c r="C12" s="3"/>
    </row>
    <row r="14" spans="2:3">
      <c r="B14" s="1">
        <f>1+2</f>
        <v>3</v>
      </c>
      <c r="C14" s="1"/>
    </row>
  </sheetData>
  <mergeCells count="7">
    <mergeCell ref="B14:C14"/>
    <mergeCell ref="B2:C2"/>
    <mergeCell ref="B4:C4"/>
    <mergeCell ref="B6:C6"/>
    <mergeCell ref="B8:C8"/>
    <mergeCell ref="B10:C10"/>
    <mergeCell ref="B12:C12"/>
  </mergeCells>
  <hyperlinks>
    <hyperlink ref="B12" r:id="rId1"/>
  </hyperlinks>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
