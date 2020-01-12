###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_range);

use Test::More tests => 9;

###############################################################################
#
# Tests setup.
#
my $got;
my $expected;
my $caption;
my $cell;
my @range;


###############################################################################
#
# Test the xl_range() method.
#
@range    = ( 0, 3, 0, 1 );
$expected = 'A1:B4';
$caption  = " \tUtility: xl_range( @range ) -> $expected";
$got      = xl_range( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range() method.
#
@range    = ( 0, 9, 0, 0 );
$expected = 'A1:A10';
$caption  = " \tUtility: xl_range( @range ) -> $expected";
$got      = xl_range( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range() method.
#
@range    = ( 6, 65, 1, 1 );
$expected = 'B7:B66';
$caption  = " \tUtility: xl_range( @range ) -> $expected";
$got      = xl_range( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range() method.
#
@range    = ( 1, 8, 2, 2 );
$expected = 'C2:C9';
$caption  = " \tUtility: xl_range( @range ) -> $expected";
$got      = xl_range( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range() method. With absolute values.
#
@range    = ( 0, 3, 0, 4, 1, 0, 0, 0 );
$expected = 'A$1:E4';
$caption  = " \tUtility: xl_range( @range ) -> $expected";
$got      = xl_range( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range() method. With absolute values.
#
@range    = ( 0, 3, 0, 4, 0, 1, 0, 0 );
$expected = 'A1:E$4';
$caption  = " \tUtility: xl_range( @range ) -> $expected";
$got      = xl_range( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range() method. With absolute values.
#
@range    = ( 0, 3, 0, 4, 0, 0, 1, 0 );
$expected = '$A1:E4';
$caption  = " \tUtility: xl_range( @range ) -> $expected";
$got      = xl_range( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range() method. With absolute values.
#
@range    = ( 0, 3, 0, 4, 0, 0, 0, 1 );
$expected = 'A1:$E4';
$caption  = " \tUtility: xl_range( @range ) -> $expected";
$got      = xl_range( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range() method. With absolute values.
#
@range    = ( 0, 3, 0, 4, 1, 1, 1, 1 );
$expected = '$A$1:$E$4';
$caption  = " \tUtility: xl_range( @range ) -> $expected";
$got      = xl_range( @range );
is( $got, $expected, $caption );


__END__


