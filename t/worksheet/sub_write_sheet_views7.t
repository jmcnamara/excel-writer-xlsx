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

use Test::More tests => 6;


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
# Freeze panes with shifted top left cell.
#


###############################################################################
#
# 1. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A21" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A2' );
$worksheet->freeze_panes( 1 , 0, 20, 0 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A21" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A1' );
$worksheet->freeze_panes( 1 , 0, 20, 0 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="E1" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'B1' );
$worksheet->freeze_panes( 0 , 1, 0, 4 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="E1" activePane="topRight" state="frozen"/><selection pane="topRight"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A1' );
$worksheet->freeze_panes( 0 , 1, 0, 4 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="I7" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="G4" sqref="G4"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'G4' );
$worksheet->freeze_panes( 3, 6, 6, 8 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="I7" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A1' );
$worksheet->freeze_panes( 3, 6, 6, 8 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );



__END__
