###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_col_to_name);

use Test::More tests => 3;

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
    push @$expected, [ $i, $cell ];
    $cell++;
}

$cell = 'WQK';
for my $i ( 16_000 .. 16_384 ) {
    push @$expected, [ $i, $cell ];
    $cell++;
}


###############################################################################
#
# Test the xl_col_to_name method.
#
$caption = " \tUtility: xl_col_to_name()";

for my $aref ( @$expected ) {
    push @$got, [ $aref->[0], xl_col_to_name( $aref->[0] ) ];
}

is_deeply( $got, $expected, $caption );


###############################################################################
#
# Test the xl_col_to_name method with absolute references.
#
$expected = '$A';
$got = xl_col_to_name( 0, 1 );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_col_to_name method for the Pod example.
#
$expected = 'AAA';
$got = xl_col_to_name( 702 );
is( $got, $expected, $caption );


__END__


