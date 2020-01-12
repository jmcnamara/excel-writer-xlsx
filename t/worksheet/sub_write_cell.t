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

use Test::More tests => 5;

###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;
my $format = undef;

###############################################################################
#
# Test the _write_cell() method for numbers.
#
$caption  = " \tWorksheet: _write_cell()";
$expected = '<c r="A1"><v>1</v></c>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_cell( 0, 0, [ 'n', 1 ] );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_cell() method for strings.
#
$caption  = " \tWorksheet: _write_cell()";
$expected = '<c r="B4" t="s"><v>0</v></c>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_cell( 3, 1, [ 's', 0 ] );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_cell() method for formulas with an optional value.
#
$caption  = " \tWorksheet: _write_cell()";
$expected = '<c r="C2"><f>A3+A5</f><v>0</v></c>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_cell( 1, 2, [ 'f', 'A3+A5', $format, 0 ] );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_cell() method for formulas without an optional value.
#
$caption  = " \tWorksheet: _write_cell()";
$expected = '<c r="C2"><f>A3+A5</f><v>0</v></c>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_cell( 1, 2, [ 'f', 'A3+A5'] );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_cell() method for array formulas with an optional value.
#
$caption  = " \tWorksheet: _write_cell()";
$expected = '<c r="A1"><f t="array" ref="A1">SUM(B1:C1*B2:C2)</f><v>9500</v></c>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_cell( 0, 0, [ 'a', 'SUM(B1:C1*B2:C2)', $format, 'A1', 9500 ] );

is( $got, $expected, $caption );


__END__


