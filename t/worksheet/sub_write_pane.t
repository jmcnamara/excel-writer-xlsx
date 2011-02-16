###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# reverse('©'), February 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 10;


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
$expected = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" />';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 1 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen" />';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 0, 1 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen" />';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 1, 1 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane ySplit="1" topLeftCell="A20" activePane="bottomLeft" state="frozen" />';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 1, 0, 19 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen" />';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 'G4' );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_panes() method.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozenSplit" />';

$worksheet = _new_worksheet(\$got);

$worksheet->freeze_panes( 3, 6, 3, 6, 1);
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 7. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane ySplit="600" activePane="bottomLeft" />';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 15 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 8. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="1280" />';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 0, 8.43 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 9. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="1280" />';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 0, 8.43 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 10. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane ySplit="600" topLeftCell="A20" activePane="bottomLeft" />';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 15, 0, 19 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


###############################################################################
#
# 10. Test the _write_panes() method. Split panes.
#
$caption  = " \tWorksheet: _write_panes()";
$expected = '<pane xSplit="6150" ySplit="1200" />';

$worksheet = _new_worksheet(\$got);

$worksheet->split_panes( 45, 42.15 );
$worksheet->_write_panes();

is( $got, $expected, $caption );


__END__


