###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_parse_time);

use Test::More tests => 14;


###############################################################################
#
# Tests setup.
#
my $got;
my $expected;
my $caption;
my $time;


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '12';
$expected = undef;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '12:18';
$expected = .5125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '12:18:00';
$expected = .5125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '12:18:00 PM';
$expected = .5125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '12:18:00 pm';
$expected = .5125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '00:18';
$expected = .0125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '00:18:00';
$expected = .0125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '00:18 AM';
$expected = .0125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '00:18:00 am';
$expected = .0125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '12:43:12';
$expected = .53;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method.
#
$time     = '14:24:00';
$expected = .60;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method. Parse hours > 24. Fix for issue #11.
#
$time     = '36:18:00';
$expected = 1.5125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method. Parse hours > 24. Fix for issue #11.
#
$time     = '108:18:00';
$expected = 4.5125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_parse_time() method. Parse hours > 24. Fix for issue #11.
#
$time     = '1068:18:00';
$expected = 44.5125;
$caption  = " \tUtility: xl_parse_time( $time )";
$got      = xl_parse_time( $time );
is( $got, $expected, $caption );




__END__


