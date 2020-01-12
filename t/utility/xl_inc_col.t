###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_inc_col);

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
# Test the xl_inc_col() method.
#
$cell     = 'A1';
$expected = 'B1';
$caption  = " \tUtility: xl_inc_col( $cell ) -> $expected";
$got      = xl_inc_col( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_inc_col() method.
#
$cell     = 'Z1';
$expected = 'AA1';
$caption  = " \tUtility: xl_inc_col( $cell ) -> $expected";
$got      = xl_inc_col( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_inc_col() method.
#
$cell     = '$B1';
$expected = '$C1';
$caption  = " \tUtility: xl_inc_col( $cell ) -> $expected";
$got      = xl_inc_col( $cell );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_inc_col() method.
#
$cell     = '$D$5';
$expected = '$E$5';
$caption  = " \tUtility: xl_inc_col( $cell ) -> $expected";
$got      = xl_inc_col( $cell );
is( $got, $expected, $caption );


__END__


