###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2026, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions qw(_new_worksheet _new_workbook);
use strict;
use warnings;

use Test::More tests => 28;


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


###############################################################################
#
# Test writing strings at Excel's length limit in UTF-16 units. Characters
# above U+FFFF, such as emojis, count as 2.
#
$caption = " \tWorksheet: write_string() UTF-16 length.";

my $workbook = _new_workbook( \$tmp );
$worksheet = $workbook->add_worksheet();

$got = $worksheet->write_string( 0, 0, "\x{1F600}" x 16383 . 'a' );
is( $got, 0, $caption );

$got = $worksheet->write_string( 1, 0, "\x{1F600}" x 16384 );
is( $got, -3, $caption );

# The string is truncated without splitting an emoji.
ok( exists ${ $worksheet->{_str_table} }->{ "\x{1F600}" x 16383 }, $caption );


###############################################################################
#
# Test the write_url() tool tip limit of 255 UTF-16 units. Characters above
# U+FFFF, such as emojis, count as 2.
#
$caption = " \tWorksheet: write_url() tool tip UTF-16 length.";

{
    local $SIG{__WARN__} = sub { };

    $got = $worksheet->write_url( 2, 0, 'https://www.perl.org/', undef, undef,
        "\x{1F600}" x 127 . 'a' );
    is( $got, 0, $caption );

    $got = $worksheet->write_url( 3, 0, 'https://www.perl.org/', undef, undef,
        "\x{1F600}" x 128 );
    is( $got, -4, $caption );
}


__END__
