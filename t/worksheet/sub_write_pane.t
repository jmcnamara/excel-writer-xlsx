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

use Test::More tests => 14;


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
# 1. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 1 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 0, 1 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 1, 1 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane ySplit="1" topLeftCell="A20" activePane="bottomLeft" state="frozen"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 1, 0, 19 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 'G4' );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozenSplit"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 3, 6, 3, 6, 1);
$worksheet->_write_panes();

is( $got, $expected, $caption );


# Split panes tests.


###############################################################################
#
# 7. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane ySplit="600" topLeftCell="A2"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 15 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 8. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane ySplit="900" topLeftCell="A3"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 30 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 9. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane ySplit="2400" topLeftCell="A8"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 105 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 10. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="1350" topLeftCell="B1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 0, 8.43 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 11. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="2310" topLeftCell="C1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 0, 17.57 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 12. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="5190" topLeftCell="F1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 0, 45 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 13. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="1350" ySplit="600" topLeftCell="B2"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 15, 8.43 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 14. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="6150" ySplit="1200" topLeftCell="G4"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 45, 54.14 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


__END__


