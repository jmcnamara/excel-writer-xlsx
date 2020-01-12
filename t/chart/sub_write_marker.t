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
# Test the _write_marker() method.
#
$caption  = " \tChart: _write_marker()";
$expected = '<c:marker><c:symbol val="none"/></c:marker>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->{_default_marker} = { type => 'none' };

$chart->_write_marker();

is( $got, $expected, $caption );

__END__


