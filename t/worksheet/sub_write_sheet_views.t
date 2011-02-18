###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# reverse('�'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 32;


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
# 1. Test the _write_sheet_views() method.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0" /></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


# Freeze panes.


###############################################################################
#
# 2. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->freeze_panes( 1 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen" /><selection pane="topRight" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->freeze_panes( 0, 1 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen" /><selection pane="topRight" activeCell="B1" sqref="B1" /><selection pane="bottomLeft" activeCell="A2" sqref="A2" /><selection pane="bottomRight" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->freeze_panes( 1, 1 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen" /><selection pane="topRight" activeCell="G1" sqref="G1" /><selection pane="bottomLeft" activeCell="A4" sqref="A4" /><selection pane="bottomRight" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->freeze_panes( 'G4' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_sheet_views() method with freeze panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozenSplit" /><selection pane="topRight" activeCell="G1" sqref="G1" /><selection pane="bottomLeft" activeCell="A4" sqref="A4" /><selection pane="bottomRight" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->freeze_panes(  3, 6, 3, 6, 1 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


# Split panes tests.


###############################################################################
#
# 7. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="600" topLeftCell="A2" /><selection pane="bottomLeft" activeCell="A2" sqref="A2" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 15 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 8. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="900" topLeftCell="A3" /><selection pane="bottomLeft" activeCell="A3" sqref="A3" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 30 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 9. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="2400" topLeftCell="A8" /><selection pane="bottomLeft" activeCell="A8" sqref="A8" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 105 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 10. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" topLeftCell="B1" /><selection pane="topRight" activeCell="B1" sqref="B1" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 0, 8.43 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 11. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="2310" topLeftCell="C1" /><selection pane="topRight" activeCell="C1" sqref="C1" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 0, 17.57 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 12. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="5190" topLeftCell="F1" /><selection pane="topRight" activeCell="F1" sqref="F1" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 0, 45 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 13. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" ySplit="600" topLeftCell="B2" /><selection pane="topRight" activeCell="B1" sqref="B1" /><selection pane="bottomLeft" activeCell="A2" sqref="A2" /><selection pane="bottomRight" activeCell="B2" sqref="B2" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 15, 8.43 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 14. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6150" ySplit="1200" topLeftCell="G4" /><selection pane="topRight" activeCell="G1" sqref="G1" /><selection pane="bottomLeft" activeCell="A4" sqref="A4" /><selection pane="bottomRight" activeCell="G4" sqref="G4" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 45, 54.14 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# Repeat of tests 7-14 with explicit topLeft cells.
#

###############################################################################
#
# 15. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="600" topLeftCell="A2" /><selection pane="bottomLeft" activeCell="A2" sqref="A2" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 15, 0, 1, 0 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 16. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="900" topLeftCell="A3" /><selection pane="bottomLeft" activeCell="A3" sqref="A3" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 30, 0, 2, 0 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 17. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="2400" topLeftCell="A8" /><selection pane="bottomLeft" activeCell="A8" sqref="A8" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 105, 0, 7, 0 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 18. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" topLeftCell="B1" /><selection pane="topRight" activeCell="B1" sqref="B1" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 0, 8.43, 0, 1 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 19. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="2310" topLeftCell="C1" /><selection pane="topRight" activeCell="C1" sqref="C1" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 0, 17.57, 0, 2 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 20. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="5190" topLeftCell="F1" /><selection pane="topRight" activeCell="F1" sqref="F1" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 0, 45, 0, 5 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 21. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" ySplit="600" topLeftCell="B2" /><selection pane="topRight" activeCell="B1" sqref="B1" /><selection pane="bottomLeft" activeCell="A2" sqref="A2" /><selection pane="bottomRight" activeCell="B2" sqref="B2" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 15, 8.43, 1, 1 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 22. Test the _write_sheet_views() method with split panes.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6150" ySplit="1200" topLeftCell="G4" /><selection pane="topRight" activeCell="G1" sqref="G1" /><selection pane="bottomLeft" activeCell="A4" sqref="A4" /><selection pane="bottomRight" activeCell="G4" sqref="G4" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->split_panes( 45, 54.14, 3, 6 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


# Test for set_selection().


###############################################################################
#
# 23. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0" /></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A1' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 24. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A2" sqref="A2" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A2' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 25. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'B1' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 26. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'D3' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 27. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3:F4" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'D3:F4' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 28. Test the _write_sheet_views() method with selection set.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="F4" sqref="D3:F4" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'F4:D3' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


# Freeze panes with selection.


###############################################################################
#
# 29. Test the _write_sheet_views() method with freeze panes + selection.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" activeCell="A2" sqref="A2" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'A2' );
$worksheet->freeze_panes( 1, 0 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 30. Test the _write_sheet_views() method with freeze panes + selection.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen" /><selection pane="topRight" activeCell="B1" sqref="B1" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'B1' );
$worksheet->freeze_panes( 0, 1 );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 31. Test the _write_sheet_views() method with freeze panes + selection.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen" /><selection pane="topRight" activeCell="G1" sqref="G1" /><selection pane="bottomLeft" activeCell="A4" sqref="A4" /><selection pane="bottomRight" activeCell="G4" sqref="G4" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'G4' );
$worksheet->freeze_panes( 'G4' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


###############################################################################
#
# 32. Test the _write_sheet_views() method with freeze panes + selection.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen" /><selection pane="topRight" activeCell="G1" sqref="G1" /><selection pane="bottomLeft" activeCell="A4" sqref="A4" /><selection pane="bottomRight" activeCell="I5" sqref="I5" /></sheetView></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->set_selection( 'I5' );
$worksheet->freeze_panes( 'G4' );
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );



__END__
