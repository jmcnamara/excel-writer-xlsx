###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_range_formula);

use Test::More tests => 6;

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
# Test the xl_range_formula() method.
#
@range    = ( 'Sheet1', 0, 9, 0, 0 );
$expected = '=Sheet1!$A$1:$A$10';
$caption  = " \tUtility: xl_range_formula( @range ) -> $expected";
$got      = xl_range_formula( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range_formula() method.
#
@range    = ( 'Sheet2',   6, 65, 1, 1 );
$expected = '=Sheet2!$B$7:$B$66';
$caption  = " \tUtility: xl_range_formula( @range ) -> $expected";
$got      = xl_range_formula( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range_formula() method.
#
@range    = ( 'New data', 1,  8, 2, 2 );
$expected = q(='New data'!$C$2:$C$9);
$caption  = " \tUtility: xl_range_formula( @range ) -> $expected";
$got      = xl_range_formula( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range_formula() method.
#
@range    = ( q('New data'), 1,  8, 2, 2 );
$expected = q(='New data'!$C$2:$C$9);
$caption  = " \tUtility: xl_range_formula( @range ) -> $expected";
$got      = xl_range_formula( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range_formula() method.
#
@range    = ( 'Sheet1', 1, 9, 0, 0 );
$expected = '=Sheet1!$A$2:$A$10';
$caption  = " \tUtility: xl_range_formula( @range ) -> $expected";
$got      = xl_range_formula( @range );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_range_formula() method.
#
@range    = ( 'Sheet1', 1, 9, 1, 1 );
$expected = '=Sheet1!$B$2:$B$10';
$caption  = " \tUtility: xl_range_formula( @range ) -> $expected";
$got      = xl_range_formula( @range );
is( $got, $expected, $caption );


__END__


