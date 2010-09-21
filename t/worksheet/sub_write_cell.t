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
# Test the _write_cell() method for TODO.
#
$caption  = " \tWorksheet: _write_cell()";
$expected = '<c r="A1"><v>1</v></c>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_cell( 0, 0, [ 'n', 1 ] );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_cell() method for TODO.
#
$caption  = " \tWorksheet: _write_cell()";
$expected = '<c r="B4" t="s"><v>0</v></c>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_cell( 3, 1, [ 's', 0 ] );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_cell() method for TODO.
#
$caption  = " \tWorksheet: _write_cell()";
$expected = '<c r="C2"><f>A3+A5</f><v>0</v></c>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_cell( 1, 2, [ 'f', 'A3+A5', undef, 0 ] );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_cell() method for TODO.
#
$caption  = " \tWorksheet: _write_cell()";
$expected = '<c r="C2"><f>A3+A5</f><v></v></c>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_cell( 1, 2, [ 'f', 'A3+A5'] );

is( $got, $expected, $caption );


__END__


