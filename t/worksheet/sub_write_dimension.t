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

use Test::More tests => 10;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;
my $cell_ref;


###############################################################################
#
# 1. Test the _write_dimension() method with no dimensions set.
#
$caption  = " \tWorksheet: _write_dimension(undef)";
$expected = '<dimension ref="A1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref"/>);

$worksheet = _new_worksheet(\$got);
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1048576';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref"/>);

$worksheet = _new_worksheet(\$got);
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'XFD1';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref"/>);

$worksheet = _new_worksheet(\$got);
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'XFD1048576';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref"/>);

$worksheet = _new_worksheet(\$got);
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref"/>);

$worksheet = _new_worksheet(\$got);
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 7. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1:B2';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref"/>);

$worksheet = _new_worksheet(\$got);
$worksheet->write( 'A1', 'some string' );
$worksheet->write( 'B2', 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 8. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1:B2';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref"/>);

$worksheet = _new_worksheet(\$got);
$worksheet->write( 'B2', 'some string' );
$worksheet->write( 'A1', 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 9. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'B2:H11';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref"/>);

$worksheet = _new_worksheet(\$got);
$worksheet->write( 'B2',  'some string' );
$worksheet->write( 'H11', 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 10. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1:XFD1048576';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref"/>);

$worksheet = _new_worksheet(\$got);
$worksheet->write( 'A1',         'some string' );
$worksheet->write( 'XFD1048576', 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


__END__
