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

use Test::More tests => 4;


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
# Test the _write_style() method.
#
$caption  = " \tChart: _write_style()";
$expected = '<c:style val="1"/>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_style( 1 );

$chart->_write_style();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_style() method.
#
$caption  = " \tChart: _write_style()";
$expected = '';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

# Default style. Not written.
$chart->set_style( 2 );

$chart->_write_style();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_style() method.
#
$caption  = " \tChart: _write_style()";
$expected = '';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

# Outside style range.
$chart->set_style( -1 );

$chart->_write_style();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_style() method.
#
$caption  = " \tChart: _write_style()";
$expected = '';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

# Outside style range.
$chart->set_style( 49 );

$chart->_write_style();

is( $got, $expected, $caption );




__END__


