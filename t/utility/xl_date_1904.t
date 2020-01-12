###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_date_1904);

use Test::More tests => 4;


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
# Test the xl_date_1904() method.
#
$date     = 0;
$expected = 0;
$caption  = " \tUtility: xl_date_1904( $date )";
$got      = xl_date_1904( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_1904() method.
#
$date     = 1462;
$expected = 0;
$caption  = " \tUtility: xl_date_1904( $date )";
$got      = xl_date_1904( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_1904() method.
#
$date     = 36526;
$expected = 35064;
$caption  = " \tUtility: xl_date_1904( $date )";
$got      = xl_date_1904( $date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_1904() method.
#
$date     = 41255;
$expected = 39793;
$caption  = " \tUtility: xl_date_1904( $date )";
$got      = xl_date_1904( $date );
is( $got, $expected, $caption );

__END__


