###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 11;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_margins(0.5);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_margins_LR(0.5);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.7" right="0.7" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_margins_TB(0.5);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.5" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_margin_left(0.5);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.7" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_margin_right(0.5);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.7" right="0.7" top="0.5" bottom="0.75" header="0.3" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_margin_top(0.5);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.5" header="0.3" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_margin_bottom(0.5);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.5" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_header('', 0.5);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.5"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_footer('', 0.5);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>';

$worksheet = _new_worksheet(\$got);

# Test numeric value with whitespace.
$worksheet->set_margins( " 0.5\n");

$worksheet->_write_page_margins();

is( $got, $expected, $caption );



__END__


