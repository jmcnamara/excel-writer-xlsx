###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_dec_col);

use Test::More tests => 4;

###############################################################################
#
# Tests setup.
#
my $got;
my $expected;
my $caption;
my $cell;


###############################################################################
#
# Test the xl_dec_col() method.
#
$cell     = 'B1';
$expected = 'A1';
$caption  = " \tUtility: xl_dec_col( $cell ) -> $expected";
$got      = xl_dec_col( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_dec_col() method.
#
$cell     = 'AA1';
$expected = 'Z1';
$caption  = " \tUtility: xl_dec_col( $cell ) -> $expected";
$got      = xl_dec_col( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_dec_col() method.
#
$cell     = '$C1';
$expected = '$B1';
$caption  = " \tUtility: xl_dec_col( $cell ) -> $expected";
$got      = xl_dec_col( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_dec_col() method.
#
$cell     = '$E$5';
$expected = '$D$5';
$caption  = " \tUtility: xl_dec_col( $cell ) -> $expected";
$got      = xl_dec_col( $cell );
is( $got, $expected, $caption );


__END__


