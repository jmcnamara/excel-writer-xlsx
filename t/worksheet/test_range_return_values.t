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

use Test::More tests => 23;


###############################################################################
#
# Tests setup.
#
my $got;
my $expected  = -2;
my $caption   = " \tWorksheet: test range return values.";
my $tmp       = '';
my $worksheet = _new_worksheet( \$tmp );
my $format    = Excel::Writer::XLSX::Format->new( {}, {}, xf_index => 1 );
my $max_row   = 1048576;
my $max_col   = 16384;

$got = $worksheet->write_string( $max_row, 0, 'Foo' );
is( $got, $expected, $caption );

$got = $worksheet->write_string( 0, $max_col, 'Foo' );
is( $got, $expected, $caption );

$got = $worksheet->write_string( $max_row, $max_col, 'Foo' );
is( $got, $expected, $caption );

$got = $worksheet->write_number( $max_row, 0, 123 );
is( $got, $expected, $caption );

$got = $worksheet->write_number( 0, $max_col, 123 );
is( $got, $expected, $caption );

$got = $worksheet->write_number( $max_row, $max_col, 123 );
is( $got, $expected, $caption );

$got = $worksheet->write_blank( $max_row, 0, $format );
is( $got, $expected, $caption );

$got = $worksheet->write_blank( 0, $max_col, $format );
is( $got, $expected, $caption );

$got = $worksheet->write_blank( $max_row, $max_col, $format );
is( $got, $expected, $caption );

$got = $worksheet->write_formula( $max_row, 0, '=A1' );
is( $got, $expected, $caption );

$got = $worksheet->write_formula( 0, $max_col, '=A1' );
is( $got, $expected, $caption );

$got = $worksheet->write_formula( $max_row, $max_col, '=A1' );
is( $got, $expected, $caption );

$got = $worksheet->write_array_formula( 0, 0, 0, $max_col, '=A1' );
is( $got, $expected, $caption );

$got = $worksheet->write_array_formula( 0, 0, $max_row, 0, '=A1' );
is( $got, $expected, $caption );

$got = $worksheet->write_array_formula( 0, $max_col, 0, 0, '=A1' );
is( $got, $expected, $caption );

$got = $worksheet->write_array_formula( $max_row, 0, 0, 0, '=A1' );
is( $got, $expected, $caption );

$got = $worksheet->write_array_formula( $max_row, $max_col, $max_row, $max_col,
    '=A1' );
is( $got, $expected, $caption );

$got = $worksheet->merge_range( 0, 0, 0, $max_col, 'Foo', $format );
is( $got, $expected, $caption );

$got = $worksheet->merge_range( 0, 0, $max_row, 0, 'Foo', $format );
is( $got, $expected, $caption );

$got = $worksheet->merge_range( 0, $max_col, 0, 0, 'Foo', $format );
is( $got, $expected, $caption );

$got = $worksheet->merge_range( $max_row, 0, 0, 0, 'Foo', $format );
is( $got, $expected, $caption );

# Column out of bounds.
$got = $worksheet->set_column( 6, $max_col, 17 );
is( $got, $expected, $caption );

$got = $worksheet->set_column( $max_col, 6, 17 );
is( $got, $expected, $caption );


__END__
