###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_workbook);
use strict;
use warnings;

use Test::More tests => 7;


###############################################################################
#
# Tests setup.
#
my @expected;
my @got;
my $formula;
my $tmp_str = '';
my $caption;
my $workbook;


###############################################################################
#
# Test the _get_chart_range() method. Simple formula.
#
$caption  = " \tWorkbook: _get_chart_range()";
$formula  = q(Sheet1!$B$1:$B$5);
@expected = ( 'Sheet1', 0, 1, 4, 1 );

$workbook = _new_workbook( \$tmp_str );

@got = $workbook->_get_chart_range( $formula );

is_deeply( \@got, \@expected, $caption );


###############################################################################
#
# Test the _get_chart_range() method. Sheet name with space.
#
$caption  = " \tWorkbook: _get_chart_range()";
$formula  = q('Sheet 1'!$B$1:$B$5);
@expected = ( q(Sheet 1), 0, 1, 4, 1 );

$workbook = _new_workbook( \$tmp_str );

@got = $workbook->_get_chart_range( $formula );

is_deeply( \@got, \@expected, $caption );


###############################################################################
#
# Test the _get_chart_range() method. Singe cell range.
#
$caption  = " \tWorkbook: _get_chart_range()";
$formula  = q(Sheet1!$B$1);
@expected = ( 'Sheet1', 0, 1, 0, 1 );

$workbook = _new_workbook( \$tmp_str );

@got = $workbook->_get_chart_range( $formula );

is_deeply( \@got, \@expected, $caption );


###############################################################################
#
# Test the _get_chart_range() method. Sheet name with an apostrophe.
#
$caption  = " \tWorkbook: _get_chart_range()";
$formula  = q('Don''t'!$B$1:$B$5);
@expected = ( q(Don't), 0, 1, 4, 1 );

$workbook = _new_workbook( \$tmp_str );

@got = $workbook->_get_chart_range( $formula );

is_deeply( \@got, \@expected, $caption );


###############################################################################
#
# Test the _get_chart_range() method. Sheet name with exclamation mark.
#
$caption  = " \tWorkbook: _get_chart_range()";
$formula  = q('aa!bb'!$B$1:$B$5);
@expected = ( q(aa!bb), 0, 1, 4, 1 );

$workbook = _new_workbook( \$tmp_str );

@got = $workbook->_get_chart_range( $formula );

is_deeply( \@got, \@expected, $caption );


###############################################################################
#
# Test the _get_chart_range() method. Invalid range.
#
$caption  = " \tWorkbook: _get_chart_range()";
$formula  = '';
@expected = ( undef );

$workbook = _new_workbook( \$tmp_str );

@got = $workbook->_get_chart_range( $formula );

is_deeply( \@got, \@expected, $caption );


###############################################################################
#
# Test the _get_chart_range() method. Invalid 2D range.
#
$caption  = " \tWorkbook: _get_chart_range()";
$formula  = q(Sheet1!$B$1:$F$5);
@expected = ( undef );

$workbook = _new_workbook( \$tmp_str );

@got = $workbook->_get_chart_range( $formula );

is_deeply( \@got, \@expected, $caption );


__END__


