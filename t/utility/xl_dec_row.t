###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_dec_row);

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
# Test the xl_dec_row() method.
#
$cell     = 'A2';
$expected = 'A1';
$caption  = " \tUtility: xl_dec_row( $cell ) -> $expected";
$got      = xl_dec_row( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_dec_row() method.
#
$cell     = 'B$3';
$expected = 'B$2';
$caption  = " \tUtility: xl_dec_row( $cell ) -> $expected";
$got      = xl_dec_row( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_dec_row() method.
#
$cell     = '$C4';
$expected = '$C3';
$caption  = " \tUtility: xl_dec_row( $cell ) -> $expected";
$got      = xl_dec_row( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_dec_row() method.
#
$cell     = '$D$5';
$expected = '$D$4';
$caption  = " \tUtility: xl_dec_row( $cell ) -> $expected";
$got      = xl_dec_row( $cell );
is( $got, $expected, $caption );


__END__


