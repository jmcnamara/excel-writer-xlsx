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

use Test::More tests => 18;

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
        undef,
        [],
    ],

    [
        '',
        [],
    ],

    [
        '0 <  2000',
        [0, '<', 2000],
    ],

    [
        'x <  2000',
        ['x', '<', 2000],
    ],

    [
        'x >  2000',
        ['x', '>', 2000],
    ],

    [
        'x == 2000',
        ['x', '==', 2000],
    ],

    [
        'x >  2000 and x <  5000',
        ['x', '>',  2000, 'and', 'x', '<', 5000],
    ],

    [
        'x = "foo"',
        ['x', '=', 'foo'],
    ],

    [
        'x = foo',
        ['x', '=', 'foo'],
    ],

    [
        'x = "foo bar"',
        ['x', '=', 'foo bar'],
    ],

    [
        'x = "foo "" bar"',
        ['x', '=', 'foo " bar'],
    ],

    [
        'x = "foo bar" or x = "bar foo"',
        ['x', '=', 'foo bar', 'or', 'x', '=', 'bar foo'],
    ],

    [
        'x = "foo "" bar" or x = "bar "" foo"',
        ['x', '=', 'foo " bar', 'or', 'x', '=', 'bar " foo'],
    ],

    [
        'x = """"""""',
        ['x', '=', '"""'],
    ],

    [
        'x = Blanks',
        ['x', '=', 'Blanks'],
    ],

    [
        'x = NonBlanks',
        ['x', '=', 'NonBlanks'],
    ],

    [
        'top 10 %',
        ['top', 10, '%'],
    ],

    [
        'top 10 items',
        ['top', 10, 'items'],
    ],

);


###############################################################################
#
# Run the test cases.
#
for my $aref ( @tests ) {
    my $expression = $aref->[0];
    my $expected   = $aref->[1];
    my @results    = $worksheet->_extract_filter_tokens( $expression );

    my $testname = $expression || 'none';

    _is_deep_diff( \@results, $expected, " \t" . $testname );
}

__END__
