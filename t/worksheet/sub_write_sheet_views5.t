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
# Test for set_selection().
#

###############################################################################
#
# 1. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A1' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A2" sqref="A2"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A2' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'B1' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'D3' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3:F4"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'D3:F4' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="F4" sqref="D3:F4"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'F4:D3' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 7. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A2" sqref="A2"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A2:A2' ); # Should be the same as 'A2'
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


__END__
