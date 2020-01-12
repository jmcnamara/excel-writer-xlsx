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

use Test::More tests => 8;


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
# Test the _write_print_options() method. Without any options.
#
$caption  = " \tWorksheet: _write_print_options()";
$expected = undef;

$worksheet = _new_worksheet(\$got);

$worksheet->_write_print_options();

is( $got, $expected, $caption );
$got = ''; # Reset after previous undef value;


###############################################################################
#
# Test the _write_print_options() method.
#
$caption  = " \tWorksheet: _write_print_options()";
$expected = '<printOptions horizontalCentered="1"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->center_horizontally();

$worksheet->_write_print_options();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_print_options() method.
#
$caption  = " \tWorksheet: _write_print_options()";
$expected = '<printOptions verticalCentered="1"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->center_vertically();

$worksheet->_write_print_options();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_print_options() method.
#
$caption  = " \tWorksheet: _write_print_options()";
$expected = '<printOptions horizontalCentered="1" verticalCentered="1"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->center_horizontally();
$worksheet->center_vertically();

$worksheet->_write_print_options();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_print_options() method.
#
$caption  = " \tWorksheet: _write_print_options()";
$expected = '';

$worksheet = _new_worksheet(\$got);

$worksheet->hide_gridlines();

$worksheet->_write_print_options();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_print_options() method.
#
$caption  = " \tWorksheet: _write_print_options()";
$expected = '<printOptions gridLines="1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->hide_gridlines(0);

$worksheet->_write_print_options();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_print_options() method.
#
$caption  = " \tWorksheet: _write_print_options()";
$expected = '';

$worksheet = _new_worksheet(\$got);

$worksheet->hide_gridlines();

$worksheet->_write_print_options(1);

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_print_options() method.
#
$caption  = " \tWorksheet: _write_print_options()";
$expected = '<printOptions gridLines="1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->hide_gridlines(0);

$worksheet->_write_print_options(2);

is( $got, $expected, $caption );

__END__


