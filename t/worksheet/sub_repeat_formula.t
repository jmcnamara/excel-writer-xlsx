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
my $formula;
my $row;
my $col;
my $worksheet = _new_worksheet(\$got);
my $format    = undef;


###############################################################################
#
# Test the store_formula/repeat_formula methods.
#
$caption  = " \tWorksheet: repeat_formula()";
$expected = 'SUM(A1:A10)';
$row      = 1;
$col      = 0;

$formula = $worksheet->store_formula( '=SUM(A1:A10)' );
$worksheet->repeat_formula( $row, $col, $formula, $format );

$got = $worksheet->{_table}->{$row}->{$col}->[1];

is( $got, $expected, $caption );


###############################################################################
#
# Test the store_formula/repeat_formula methods.
#
$caption  = " \tWorksheet: repeat_formula()";
$expected = 'SUM(A2:A10)';
$row      = 2;
$col      = 0;

$formula = $worksheet->store_formula( '=SUM(A1:A10)' );
$worksheet->repeat_formula( $row, $col, $formula, $format, 'A1', 'A2' );

$got = $worksheet->{_table}->{$row}->{$col}->[1];

is( $got, $expected, $caption );


###############################################################################
#
# Test the store_formula/repeat_formula methods.
#
$caption  = " \tWorksheet: repeat_formula()";
$expected = 'SUM(A2:A10)';
$row      = 3;
$col      = 0;

$formula = $worksheet->store_formula( '=SUM(A1:A10)' );
$worksheet->repeat_formula( $row, $col, $formula, $format, qr/^A1$/, 'A2' );

$got = $worksheet->{_table}->{$row}->{$col}->[1];

is( $got, $expected, $caption );


###############################################################################
#
# Test the store_formula/repeat_formula methods.
#
$caption  = " \tWorksheet: repeat_formula()";
$expected = 'A2+A2';
$row      = 4;
$col      = 0;

$formula = $worksheet->store_formula( 'A1+A1' );
$worksheet->repeat_formula( $row, $col, $formula, $format, ('A1', 'A2') x 2 );

$got = $worksheet->{_table}->{$row}->{$col}->[1];

is( $got, $expected, $caption );


###############################################################################
#
# Test the store_formula/repeat_formula methods.
#
$caption  = " \tWorksheet: repeat_formula()";
$expected = 'A10 + SIN(A10)';
$row      = 5;
$col      = 0;

$formula = $worksheet->store_formula( 'A1 + SIN(A1)' );
$worksheet->repeat_formula( $row, $col, $formula, $format, (qr/^A1$/, 'A10') x 2 );

$got = $worksheet->{_table}->{$row}->{$col}->[1];

is( $got, $expected, $caption );


__END__


