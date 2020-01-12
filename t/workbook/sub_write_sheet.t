###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_workbook';
use strict;
use warnings;

use Test::More tests => 3;

###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $workbook;


###############################################################################
#
# Test the _write_sheet() method.
#
$caption  = " \tWorkbook: _write_sheet()";
$expected = '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>';

$workbook = _new_workbook(\$got);
$workbook->_write_sheet( 'Sheet1', 1 );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_sheet() method. Hidden worksheet.
#
$caption  = " \tWorkbook: _write_sheet()";
$expected = '<sheet name="Sheet1" sheetId="1" state="hidden" r:id="rId1"/>';

$workbook = _new_workbook(\$got);
$workbook->_write_sheet( 'Sheet1', 1, 1 );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_sheet() method. Sheetname with XML character.
#
$caption  = " \tWorkbook: _write_sheet()";
$expected = '<sheet name="Bits &amp; Bobs" sheetId="1" r:id="rId1"/>';

$workbook = _new_workbook(\$got);
$workbook->_write_sheet( 'Bits & Bobs', 1 );

is( $got, $expected, $caption );

__END__


