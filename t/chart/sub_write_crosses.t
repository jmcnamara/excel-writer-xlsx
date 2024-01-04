###############################################################################
#
# Tests for Excel::Writer::XLSX::Chart methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
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
# Test the _write_crosses() method.
#
$caption  = " \tChart: _write_crosses()";
$expected = '<c:crosses val="autoZero"/>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->_write_crosses( 'autoZero' );

is( $got, $expected, $caption );

__END__


