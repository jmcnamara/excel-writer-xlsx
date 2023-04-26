###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 6;


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
# Test the _write_page_setup() method. Without any page setup.
#
$caption  = " \tWorksheet: _write_page_setup()";
$expected = undef;

$worksheet = _new_worksheet(\$got);

$worksheet->_write_page_setup();

is( $got, $expected, $caption );
$got = ''; # Reset after previous undef value;


###############################################################################
#
# Test the _write_page_setup() method. With set_landscape();
#
$caption  = " \tWorksheet: _write_page_setup()";
$expected = '<pageSetup orientation="landscape"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_landscape();

$worksheet->_write_page_setup();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_setup() method. With set_portrait();
#
$caption  = " \tWorksheet: _write_page_setup()";
$expected = '<pageSetup orientation="portrait"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_portrait();

$worksheet->_write_page_setup();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_setup() method. With set_paper();
#
$caption  = " \tWorksheet: _write_page_setup()";
$expected = '<pageSetup paperSize="9" orientation="portrait"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_paper( 9 );

$worksheet->_write_page_setup();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_setup() method. With print_across();
#
$caption  = " \tWorksheet: _write_page_setup()";
$expected = '<pageSetup pageOrder="overThenDown" orientation="portrait"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->print_across();

$worksheet->_write_page_setup();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_page_setup() method. With black_and_white();
#
$caption  = " \tWorksheet: _write_page_setup()";
$expected = '<pageSetup orientation="portrait" blackAndWhite="1"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->print_black_and_white();

$worksheet->_write_page_setup();

is( $got, $expected, $caption );


__END__


