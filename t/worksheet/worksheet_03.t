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
# Test column formatting.
#
$caption = " \tWorksheet: _assemble_xml_file()";

my $format = Excel::Writer::XLSX::Format->new( {}, {}, xf_index => 1, bold => 1 );
$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_column( 'B:D', 5 );
$worksheet->set_column( 'F:F', 8, undef, 1 );
$worksheet->set_column( 'H:H', undef, $format );
$worksheet->set_column( 'J:J', 2 );
$worksheet->set_column( 'L:L', undef, undef, 1 );

$worksheet->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="F1:H1"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="2" max="4" width="5.7109375" customWidth="1"/>
    <col min="6" max="6" width="8.7109375" hidden="1" customWidth="1"/>
    <col min="8" max="8" width="9.140625" style="1"/>
    <col min="10" max="10" width="2.7109375" customWidth="1"/>
    <col min="12" max="12" width="0" hidden="1" customWidth="1"/>
  </cols>
  <sheetData/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
