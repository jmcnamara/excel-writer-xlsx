###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 7;

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
# Test the _write_sheet_view() method. Tab not selected.
#
$caption  = " \tWorksheet: _write_sheet_view()";
$expected = '<sheetView workbookViewId="0"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->_write_sheet_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_sheet_view() method. Tab selected.
#
$caption  = " \tWorksheet: _write_sheet_view()";
$expected = '<sheetView tabSelected="1" workbookViewId="0"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->select();
$worksheet->_write_sheet_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_sheet_view() method. Tab selected + hide_gridlines().
#
$caption  = " \tWorksheet: _write_sheet_view() + hide_gridlines()";
$expected = '<sheetView tabSelected="1" workbookViewId="0"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->select();
$worksheet->hide_gridlines();
$worksheet->_write_sheet_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_sheet_view() method. Tab selected + hide_gridlines().
#
$caption  = " \tWorksheet: _write_sheet_view() + hide_gridlines(0)";
$expected = '<sheetView tabSelected="1" workbookViewId="0"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->select();
$worksheet->hide_gridlines( 0 );
$worksheet->_write_sheet_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_sheet_view() method. Tab selected + hide_gridlines().
#
$caption  = " \tWorksheet: _write_sheet_view() + hide_gridlines(1)";
$expected = '<sheetView tabSelected="1" workbookViewId="0"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->select();
$worksheet->hide_gridlines( 1 );
$worksheet->_write_sheet_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_sheet_view() method. Tab selected + hide_gridlines().
#
$caption  = " \tWorksheet: _write_sheet_view() + hide_gridlines(2)";
$expected = '<sheetView showGridLines="0" tabSelected="1" workbookViewId="0"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->select();
$worksheet->hide_gridlines( 2 );
$worksheet->_write_sheet_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_sheet_view() method. Tab selected + hide_row_col_headings().
#
$caption  = " \tWorksheet: _write_sheet_view() + hide_row_col_headings()";
$expected = '<sheetView showRowColHeaders="0" tabSelected="1" workbookViewId="0"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->select();
$worksheet->hide_row_col_headers();
$worksheet->_write_sheet_view();

is( $got, $expected, $caption );



__END__


