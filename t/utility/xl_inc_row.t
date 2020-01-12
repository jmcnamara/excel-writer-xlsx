###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_inc_row);

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
# Test the xl_inc_row() method.
#
$cell     = 'A1';
$expected = 'A2';
$caption  = " \tUtility: xl_inc_row( $cell ) -> $expected";
$got      = xl_inc_row( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_inc_row() method.
#
$cell     = 'B$2';
$expected = 'B$3';
$caption  = " \tUtility: xl_inc_row( $cell ) -> $expected";
$got      = xl_inc_row( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_inc_row() method.
#
$cell     = '$C3';
$expected = '$C4';
$caption  = " \tUtility: xl_inc_row( $cell ) -> $expected";
$got      = xl_inc_row( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_inc_row() method.
#
$cell     = '$D$4';
$expected = '$D$5';
$caption  = " \tUtility: xl_inc_row( $cell ) -> $expected";
$got      = xl_inc_row( $cell );
is( $got, $expected, $caption );


__END__


