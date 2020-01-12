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

use Test::More tests => 4;


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
# Split panes with selection.
#


###############################################################################
#
# 1. Test the _write_sheet_views() method with split panes + selection.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="600" topLeftCell="A2" activePane="bottomLeft"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A2' );
$worksheet->split_panes( 15, 0 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_sheet_views() method with split panes + selection.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" topLeftCell="B1" activePane="topRight"/><selection pane="topRight" activeCell="B1" sqref="B1"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'B1' );
$worksheet->split_panes( 0, 8.43 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_sheet_views() method with split panes + selection.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6150" ySplit="1200" topLeftCell="G4" activePane="bottomRight"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="G4" sqref="G4"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'G4' );
$worksheet->split_panes( 45, 54.14 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_sheet_views() method with split panes + selection.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6150" ySplit="1200" topLeftCell="G4" activePane="bottomRight"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="I5" sqref="I5"/></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'I5' );
$worksheet->split_panes( 45, 54.14 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );



__END__
