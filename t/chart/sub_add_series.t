###############################################################################
#
# Tests for Excel::Writer::XLSX::Chart methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_object _is_deep_diff);
use strict;
use warnings;
use Excel::Writer::XLSX::Chart;

use Test::More tests => 5;


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
# Test the add_series() method.
#
$caption  = " \tChart: add_series()";
$expected = {
    _categories    => undef,
    _values        => '=Sheet1!$A$1:$A$5',
    _name          => undef,
    _name_formula  => undef,
    _name_id       => undef,
    _cat_data_id   => undef,
    _val_data_id   => 0,
    _line          => { _defined => 0 },
    _fill          => { _defined => 0 },
    _pattern       => undef,
    _gradient      => undef,
    _marker        => undef,
    _trendline     => undef,
    _smooth        => undef,
    _error_bars    => { _x_error_bars => undef, _y_error_bars => undef },
    _points        => undef,
    _labels        => undef,
    _invert_if_neg => undef,
    _x2_axis       => undef,
    _y2_axis       => undef,
};

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->add_series( values => '=Sheet1!$A$1:$A$5' );

$got = $chart->{_series}->[0];

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test the add_series() method.
#
$caption  = " \tChart: add_series()";
$expected = [
    {
        _categories    => '=Sheet1!$A$1:$A$5',
        _values        => '=Sheet1!$B$1:$B$5',
        _name          => 'Text',
        _name_formula  => undef,
        _name_id       => undef,
        _cat_data_id   => 0,
        _val_data_id   => 1,
        _line          => { _defined => 0 },
        _fill          => { _defined => 0 },
        _gradient      => undef,
        _pattern       => undef,
        _marker        => undef,
        _trendline     => undef,
        _smooth        => undef,
        _error_bars    => { _x_error_bars => undef, _y_error_bars => undef },
        _points        => undef,
        _labels        => undef,
        _invert_if_neg => undef,
        _x2_axis       => undef,
        _y2_axis       => undef,
    }
];

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->add_series(
    categories => '=Sheet1!$A$1:$A$5',
    values     => '=Sheet1!$B$1:$B$5',
    name       => 'Text'
);

$got = $chart->{_series};

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test the add_series() method.
#
$caption  = " \tChart: add_series()";
$expected = [
    {
        _categories    => undef,
        _values        => '=Sheet1!$A$1:$A$5',
        _name          => undef,
        _name_formula  => undef,
        _name_id       => undef,
        _cat_data_id   => undef,
        _val_data_id   => 0,
        _line          => { _defined => 0 },
        _fill          => { _defined => 0 },
        _gradient      => undef,
        _pattern       => undef,
        _marker        => undef,
        _trendline     => undef,
        _smooth        => undef,
        _error_bars    => { _x_error_bars => undef, _y_error_bars => undef },
        _points        => undef,
        _labels        => undef,
        _invert_if_neg => undef,
        _x2_axis       => undef,
        _y2_axis       => undef,
    }
];

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->add_series( values => [ 'Sheet1', 0, 4, 0, 0 ] );

$got = $chart->{_series};

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test the add_series() method.
#
$caption  = " \tChart: add_series()";
$expected = {
    _categories    => '=Sheet1!$A$1:$A$5',
    _values        => '=Sheet1!$B$1:$B$5',
    _name          => 'Text',
    _name_formula  => undef,
    _name_id       => undef,
    _cat_data_id   => 0,
    _val_data_id   => 1,
    _line          => { _defined => 0 },
    _fill          => { _defined => 0 },
    _gradient      => undef,
    _pattern       => undef,
    _marker        => undef,
    _trendline     => undef,
    _smooth        => undef,
    _error_bars    => { _x_error_bars => undef, _y_error_bars => undef },
    _points        => undef,
    _labels        => undef,
    _invert_if_neg => undef,
    _x2_axis       => undef,
    _y2_axis       => undef,
};

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->add_series(
    categories => [ 'Sheet1', 0, 4, 0, 0 ],
    values     => [ 'Sheet1', 0, 4, 1, 1 ],
    name       => 'Text'
);

$got = $chart->{_series}->[0];

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test the add_series() method.
#
$caption  = " \tChart: add_series()";
$expected = {
    _categories    => '=Sheet1!$A$1:$A$5',
    _values        => '=Sheet1!$B$1:$B$5',
    _name          => 'Text',
    _name_formula  => undef,
    _name_id       => undef,
    _cat_data_id   => 0,
    _val_data_id   => 1,
    _line          => { _defined => 0 },
    _fill          => { _defined => 0 },
    _gradient      => undef,
    _pattern       => undef,
    _marker        => undef,
    _trendline     => undef,
    _smooth        => undef,
    _error_bars    => { _x_error_bars => undef, _y_error_bars => undef },
    _points        => undef,
    _labels        => undef,
    _invert_if_neg => undef,
    _x2_axis       => 1,
    _y2_axis       => 1,
};

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->add_series(
    categories => [ 'Sheet1', 0, 4, 0, 0 ],
    values     => [ 'Sheet1', 0, 4, 1, 1 ],
    name       => 'Text',
    x2_axis    => 1,
    y2_axis    => 1,
);

$got = $chart->{_series}->[0];

_is_deep_diff( $got, $expected, $caption );


__END__


