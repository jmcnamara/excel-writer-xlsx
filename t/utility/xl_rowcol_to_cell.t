###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_rowcol_to_cell);

use Test::More tests => 4;

###############################################################################
#
# Tests setup.
#
my $got;
my $expected;
my $caption;
my $cell;

# Create a test case for a range of the Excel 2007 columns.
$cell = 'A';
for my $i ( 0 .. 300 ) {
    push @$expected, [ $i, $i, $cell . ( $i + 1 ) ];
    $cell++;
}

$cell = 'WQK';
for my $i ( 16_000 .. 16_384 ) {
    push @$expected, [ $i, $i, $cell . ( $i + 1 ) ];
    $cell++;
}


###############################################################################
#
# Test the xl_rowcol_to_cell method.
#
$caption = " \tUtility: xl_rowcol_to_cell()";

for my $aref ( @$expected ) {
    push @$got,
      [ $aref->[0], $aref->[1], xl_rowcol_to_cell( $aref->[0], $aref->[1] ) ];
}

is_deeply( $got, $expected, $caption );


###############################################################################
#
# Test the xl_rowcol_to_cell method with absolute references.
#
$expected = 'A$1';
$got = xl_rowcol_to_cell( 0, 0, 1 );
is( $got, $expected, $caption );


$expected = '$A1';
$got = xl_rowcol_to_cell( 0, 0, 0, 1 );
is( $got, $expected, $caption );


$expected = '$A$1';
$got = xl_rowcol_to_cell( 0, 0, 1, 1 );
is( $got, $expected, $caption );


__END__


