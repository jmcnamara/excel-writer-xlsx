###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Tests for the token extraction method used to parse autofilter expressions.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_worksheet _is_deep_diff);
use strict;
use warnings;

use Test::More tests => 24;

###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet = _new_worksheet( \$got );


###############################################################################
#
# Test cases structured as [$input, [@expected_output]]
#
my @tests = (

    [
        'x =  2000',
        [2, 2000],
    ],

    [
        'x == 2000',
        [2, 2000],
    ],

    [
        'x =~ 2000',
        [2, 2000],
    ],

    [
        'x eq 2000',
        [2, 2000],
    ],

    [
        'x <> 2000',
        [5, 2000],
    ],

    [
        'x != 2000',
        [5, 2000],
    ],

    [
        'x ne 2000',
        [5, 2000],
    ],

    [
        'x !~ 2000',
        [5, 2000],
    ],

    [
        'x >  2000',
        [4, 2000],
    ],

    [
        'x <  2000',
        [1, 2000],
    ],

    [
        'x >= 2000',
        [6, 2000],
    ],

    [
        'x <= 2000',
        [3, 2000],
    ],

    [
        'x >  2000 and x <  5000',
        [4,  2000, 0, 1, 5000],
    ],

    [
        'x >  2000 &&  x <  5000',
        [4,  2000, 0, 1, 5000],
    ],

    [
        'x >  2000 or  x <  5000',
        [4,  2000, 1, 1, 5000],
    ],

    [
        'x >  2000 ||  x <  5000',
        [4,  2000, 1, 1, 5000],
    ],

    [
        'x =  Blanks',
        [2, 'blanks'],
    ],

    [
        'x =  NonBlanks',
        [5, ' '],
    ],

    [
        'x <> Blanks',
        [5, ' '],
    ],

    [
        'x <> NonBlanks',
        [2, 'blanks'],
    ],

    [
        'Top 10 Items',
        [30, 10],
    ],

    [
        'Top 20 %',
        [31, 20],
    ],

    [
        'Bottom 5 Items',
        [32, 5],
    ],

    [
        'Bottom 101 %',
        [33, 101],
    ],


);


###############################################################################
#
# Run the test cases.
#
for my $aref ( @tests ) {
    my $expression = $aref->[0];
    my $expected   = $aref->[1];
    my @tokens     = $worksheet->_extract_filter_tokens( $expression );
    my @results    = $worksheet->_parse_filter_expression( $expression, @tokens );

    my $testname = $expression || 'none';

    is_deeply( \@results, $expected, " \t" . $testname );
}

