###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_workbook';
use strict;
use warnings;

use Test::More tests => 4;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $workbook;


###############################################################################
#
# Test the _write_calc_pr() method.
#
$caption  = " \tWorkbook: _write_calc_pr()";
$expected = '<calcPr calcId="124519" fullCalcOnLoad="1"/>';

$workbook = _new_workbook(\$got);

$workbook->_write_calc_pr();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_calc_pr() method with the calculation mode set to
# auto_except_tables.
#
$caption  = " \tWorkbook: _write_calc_pr()";
$expected = '<calcPr calcId="124519" calcMode="autoNoTable" fullCalcOnLoad="1"/>';

$workbook = _new_workbook(\$got);
$workbook->set_calc_mode('auto_except_tables');

$workbook->_write_calc_pr();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_calc_pr() method with the calculation mode set to manual.
#
$caption  = " \tWorkbook: _write_calc_pr()";
$expected = '<calcPr calcId="124519" calcMode="manual" calcOnSave="0"/>';

$workbook = _new_workbook(\$got);
$workbook->set_calc_mode('manual');

$workbook->_write_calc_pr();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_calc_pr() method with non-default calc id.
#
$caption  = " \tWorkbook: _write_calc_pr()";
$expected = '<calcPr calcId="12345" fullCalcOnLoad="1"/>';

$workbook = _new_workbook(\$got);
$workbook->set_calc_mode('auto', 12345);

$workbook->_write_calc_pr();

is( $got, $expected, $caption );


__END__
