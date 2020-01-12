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
# Test the _write_number_format() method.
#
$caption  = " \tChart: _write_number_format()";
$expected = '<c:numFmt formatCode="General" sourceLinked="1"/>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->_write_number_format(
    {
        _num_format => 'General',
        _defaults   => { num_format => 'General' }
    }
);

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_number_format() method.
#
$caption  = " \tChart: _write_number_format()";
$expected = '<c:numFmt formatCode="#,##0.00" sourceLinked="0"/>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->_write_number_format(
    {
        _num_format => '#,##0.00',
        _defaults   => { num_format => 'General' }
    }
);

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_cat_number_format() method.
#
$caption  = " \tChart: _write_number_format()";
$expected = '';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->_write_cat_number_format(
    {
        _num_format => 'General',
        _defaults   => { num_format => 'General' }
    }
);

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_cat_number_format() method.
#
$caption  = " \tChart: _write_number_format()";
$expected = '<c:numFmt formatCode="#,##0.00" sourceLinked="0"/>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->_write_cat_number_format(
    {
        _num_format => '#,##0.00',
        _defaults   => { num_format => 'General' }
    }
);

is( $got, $expected, $caption );





__END__


