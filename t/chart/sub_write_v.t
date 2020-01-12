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
# Test the _write_v() method.
#
$caption  = " \tChart: _write_v()";
$expected = '<c:v>Apple</c:v>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->_write_v( 'Apple' );

is( $got, $expected, $caption );

__END__


