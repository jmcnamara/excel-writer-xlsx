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

use Test::More tests => 12;


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
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="r"/><c:layout/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

# Default.
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="r"/><c:layout/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

# Default.
$chart->set_legend();
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="r"/><c:layout/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'right' );
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="t"/><c:layout/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'top' );
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="l"/><c:layout/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'left' );
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="b"/><c:layout/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'bottom' );
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'none' );
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'some_non_existing_value' );
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="1"/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'overlay_right' );
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="l"/><c:layout/><c:overlay val="1"/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'overlay_left' );
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="tr"/><c:layout/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'top_right' );
$chart->_write_legend();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_legend() method.
#
$caption  = " \tChart: _write_legend()";
$expected = '<c:legend><c:legendPos val="tr"/><c:layout/><c:overlay val="1"/></c:legend>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->set_legend( position => 'overlay_top_right' );
$chart->_write_legend();

is( $got, $expected, $caption );

__END__
