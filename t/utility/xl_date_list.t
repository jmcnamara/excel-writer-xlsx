###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_date_list);

use Test::More;


eval { require Date::Calc };

if ($@) {
    plan skip_all => 'Date::Calc required to run optional tests.';
}
else {
    plan tests => 9;
}


###############################################################################
#
# Tests setup.
#
my $got;
my $expected;
my $caption;
my @date;


###############################################################################
#
# Test the xl_date_list() method.
#
@date     = ();
$expected = undef;
$caption  = " \tUtility: xl_date_list( @date )";
$got      = xl_date_list( @date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_list() method.
#
@date     = (2002, 1, 2);
$expected = 37258;
$caption  = " \tUtility: xl_date_list( @date )";
$got      = xl_date_list( @date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_list() method.
#
@date     = (2002, 1, 2, 12);
$expected = 37258.5;
$caption  = " \tUtility: xl_date_list( @date )";
$got      = xl_date_list( @date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_list() method.
#
@date     = (2002, 1, 2, 14, 24);
$expected = 37258.6;
$caption  = " \tUtility: xl_date_list( @date )";
$got      = xl_date_list( @date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_list() method.
#
@date     = (2002, 1, 2, 12, 43, 12);
$expected = 37258.53;
$caption  = " \tUtility: xl_date_list( @date )";
$got      = xl_date_list( @date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_list() method.
#
@date     = (1900, 1, 1);
$expected = 1;
$caption  = " \tUtility: xl_date_list( @date )";
$got      = xl_date_list( @date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_list() method.
#
@date     = (1900, 2, 27);
$expected = 58;
$caption  = " \tUtility: xl_date_list( @date )";
$got      = xl_date_list( @date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_list() method.
#
@date     = (1900, 2, 28);
$expected = 59;
$caption  = " \tUtility: xl_date_list( @date )";
$got      = xl_date_list( @date );
is( $got, $expected, $caption );


###############################################################################
#
# Test the xl_date_list() method.
#
@date     = (1900, 3, 1);
$expected = 61;
$caption  = " \tUtility: xl_date_list( @date )";
$got      = xl_date_list( @date );
is( $got, $expected, $caption );


__END__


