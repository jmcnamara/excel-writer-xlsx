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

use Test::More tests => 8;

###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;
my $format = Excel::Writer::XLSX::Format->new( {}, {}, xf_index => 1 );

###############################################################################
#
# 1. Test the _write_row() method.
#
$caption  = " \tWorksheet: _write_row()";
$expected = '<row r="1">';

$worksheet = _new_worksheet( \$got );

$worksheet->_write_row( 0 );

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_row() method.
#
$caption  = " \tWorksheet: _write_row()";
$expected = '<row r="3" spans="2:2">';

$worksheet = _new_worksheet( \$got );

$worksheet->_write_row( 2, '2:2' );

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_row() method.
#
$caption  = " \tWorksheet: _write_row()";
$expected = '<row r="2" ht="30" customHeight="1">';

$worksheet = _new_worksheet( \$got );

$worksheet->_write_row( 1, undef, 30 );

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_row() method.
#
$caption  = " \tWorksheet: _write_row()";
$expected = '<row r="4" hidden="1">';

$worksheet = _new_worksheet( \$got );

$worksheet->_write_row( 3, undef, undef, undef, 1 );

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_row() method.
#
$caption  = " \tWorksheet: _write_row()";
$expected = '<row r="7" s="1" customFormat="1">';

$worksheet = _new_worksheet( \$got );

$worksheet->_write_row( 6, undef, undef, $format );

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_row() method.
#
$caption  = " \tWorksheet: _write_row()";
$expected = '<row r="10" ht="3" customHeight="1">';

$worksheet = _new_worksheet( \$got );

$worksheet->_write_row( 9, undef, 3 );

is( $got, $expected, $caption );


###############################################################################
#
# 7. Test the _write_row() method.
#
$caption  = " \tWorksheet: _write_row()";
$expected = '<row r="13" ht="24" hidden="1" customHeight="1">';

$worksheet = _new_worksheet( \$got );

$worksheet->_write_row( 12, undef, 24, undef, 1 );

is( $got, $expected, $caption );


###############################################################################
#
# 8. Test the _write_empty_row() method.
#
$caption  = " \tWorksheet: _write_empty_row()";
$expected = '<row r="13" ht="24" hidden="1" customHeight="1"/>';

$worksheet = _new_worksheet( \$got );

$worksheet->_write_empty_row( 12, undef, 24, undef, 1 );

is( $got, $expected, $caption );


__END__


