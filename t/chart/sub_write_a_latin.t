###############################################################################
#
# Tests for Excel::Writer::XLSX::Chart methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_object';
use strict;
use warnings;
use Excel::Writer::XLSX::Chart;

use Test::More tests => 1;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $chart;


###############################################################################
#
# Test the _write_a_latin() method.
#
$caption  = " \tChart: _write_a_latin()";
$expected = '<a:latin typeface="Arial" pitchFamily="34" charset="0"/>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->_write_a_latin(
    'typeface'    => 'Arial',
    'pitchFamily' => 34,
    'charset'     => 0
);

is( $got, $expected, $caption );

__END__


