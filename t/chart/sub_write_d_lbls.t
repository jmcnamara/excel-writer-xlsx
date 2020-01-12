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

use Test::More tests => 18;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $chart;
my %arg;
my $labels;


###############################################################################
#
# Test the _write_d_lbls() method. Value only.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:showVal val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1 };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );



###############################################################################
#
# Test the _write_d_lbls() method. Series name only.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:showSerName val="1"/></c:dLbls>';

$arg{data_labels} = { series_name => 1 };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Category only.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:showCatName val="1"/></c:dLbls>';

$arg{data_labels} = {  category => 1 };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Value, category and series.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, category => 1, series_name => 1 };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Position = center.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="ctr"/><c:showVal val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, position => 'center' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    below       => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Position = left.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="l"/><c:showVal val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, position => 'left' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    below       => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Position = right.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="r"/><c:showVal val="1"/></c:dLbls>';


$arg{data_labels} = { value => 1, position => 'right' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    below       => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Position = above.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="t"/><c:showVal val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, position => 'above' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    below       => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Position = top.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="t"/><c:showVal val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, position => 'top' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    top         => 't',
    below       => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Position = below.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="b"/><c:showVal val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, position => 'below' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    below       => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Position = bottom.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="b"/><c:showVal val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, position => 'bottom' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    below       => 'b',
    bottom      => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Pie chart.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, leader_lines => 1 };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    below       => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Pie chart. Postion = 
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, leader_lines => 1, position => '' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    below       => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Pie chart. Postion = center.
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="ctr"/><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, leader_lines => 1, position => 'center' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->{_label_positions} = {
    center      => 'ctr',
    right       => 'r',
    left        => 'l',
    above       => 't',
    below       => 'b',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Pie chart. Postion = inside_end
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="inEnd"/><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, leader_lines => 1, position => 'inside_end' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    inside_base => 'inBase',
    inside_end  => 'inEnd',
    outside_end => 'outEnd',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Pie chart. Postion = outside_end
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="outEnd"/><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, leader_lines => 1, position => 'outside_end' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    inside_base => 'inBase',
    inside_end  => 'inEnd',
    outside_end => 'outEnd',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Pie chart. Postion = best_fit
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:dLblPos val="bestFit"/><c:showVal val="1"/><c:showLeaderLines val="1"/></c:dLbls>';

$arg{data_labels} = { value => 1, leader_lines => 1, position => 'best_fit' };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );
$chart->{_label_positions} = {
    center      => 'ctr',
    inside_base => 'inBase',
    inside_end  => 'inEnd',
    outside_end => 'outEnd',
    best_fit    => 'bestFit',
};

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_d_lbls() method. Pie chart. Percentage
#
$caption  = " \tChart: _write_d_lbls()";
$expected = '<c:dLbls><c:showPercent val="1"/><c:showLeaderLines val="1"/></c:dLbls>';

$arg{data_labels} = { leader_lines => 1, percentage => 1 };

$chart  = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$labels = $chart->_get_labels_properties( $arg{data_labels} );

$chart->_write_d_lbls( $labels );

is( $got, $expected, $caption );


__END__


