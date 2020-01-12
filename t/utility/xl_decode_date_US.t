
###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_decode_date_US);

use Test::More;


eval { require Date::Calc };

if ($@) {
    plan skip_all => 'Date::Calc required to run optional tests.';
}
else {
    plan tests => 7;
}


###############################################################################
#
# Tests setup.
#
my $got;
my $expected;
my $caption;
my $date;


###############################################################################
#
# Test the xl_decode_date_US() method.
#
$expected = undef;
$caption  = " \tUtility: xl_decode_date_US()";
$got      = xl_decode_date_US();
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_US() method.
#
$date     = '12:18';
$expected = .5125;
$caption  = " \tUtility: xl_decode_date_US( $date )";
$got      = xl_decode_date_US( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_US() method.
#
$date     = '19/19/2000';
$expected = undef;
$caption  = " \tUtility: xl_decode_date_US( $date )";
$got      = xl_decode_date_US( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_US() method.
#
$date     = '1/2/2002';
$expected = 37258;
$caption  = " \tUtility: xl_decode_date_US( $date )";
$got      = xl_decode_date_US( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_US() method.
#
$date     = '7/11/97';
$expected = 35622;
$caption  = " \tUtility: xl_decode_date_US( $date )";
$got      = xl_decode_date_US( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_US() method.
#
$date     = 'July 11 Friday 1997';
$expected = 35622;
$caption  = " \tUtility: xl_decode_date_US( $date )";
$got      = xl_decode_date_US( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_US() method.
#
$date     = '10:12 AM July 11 Friday 1997';
$expected = 35622.425;
$caption  = " \tUtility: xl_decode_date_US( $date )";
$got      = xl_decode_date_US( $date );
is( $got, $expected, $caption );


__END__


