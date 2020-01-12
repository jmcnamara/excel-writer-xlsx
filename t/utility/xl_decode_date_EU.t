###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_decode_date_EU);

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
# Test the xl_decode_date_EU() method.
#
$expected = undef;
$caption  = " \tUtility: xl_decode_date_EU()";
$got      = xl_decode_date_EU();
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_EU() method.
#
$date     = '12:18';
$expected = .5125;
$caption  = " \tUtility: xl_decode_date_EU( $date )";
$got      = xl_decode_date_EU( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_EU() method.
#
$date     = '19/19/2000';
$expected = undef;
$caption  = " \tUtility: xl_decode_date_EU( $date )";
$got      = xl_decode_date_EU( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_EU() method.
#
$date     = '2/1/2002';
$expected = 37258;
$caption  = " \tUtility: xl_decode_date_EU( $date )";
$got      = xl_decode_date_EU( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_EU() method.
#
$date     = '11/7/97';
$expected = 35622;
$caption  = " \tUtility: xl_decode_date_EU( $date )";
$got      = xl_decode_date_EU( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_EU() method.
#
$date     = 'Friday 11 July 1997';
$expected = 35622;
$caption  = " \tUtility: xl_decode_date_EU( $date )";
$got      = xl_decode_date_EU( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_decode_date_EU() method.
#
$date     = '10:12 AM Friday 11 July 1997';
$expected = 35622.425;
$caption  = " \tUtility: xl_decode_date_EU( $date )";
$got      = xl_decode_date_EU( $date );
is( $got, $expected, $caption );


__END__


