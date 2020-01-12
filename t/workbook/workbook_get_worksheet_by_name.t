###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_workbook);
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
# Test the get_worksheet_by_name() method.
#
$caption = " \tWorkbook: get_worksheet_by_name()";

$workbook = _new_workbook(\$got);

# Test a valid implicit name.
$expected = $workbook->add_worksheet();
$got      = $workbook->get_worksheet_by_name('Sheet1');
is($got, $expected, $caption);

# Test a valid explicit name.
$expected = $workbook->add_worksheet('Sheet 2');
$got      = $workbook->get_worksheet_by_name('Sheet 2');
is($got, $expected, $caption);

# Test an invalid name.
$expected = undef;
$got      = $workbook->get_worksheet_by_name('Sheet3');
is($got, $expected, $caption);

# Test an invalid name.
$expected = undef;
$got      = $workbook->get_worksheet_by_name();
is($got, $expected, $caption);



