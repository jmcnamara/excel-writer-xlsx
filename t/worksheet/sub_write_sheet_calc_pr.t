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

use Test::More tests => 1;


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
# Test the _write_sheet_calc_pr() method.
#
$caption  = " \tWorksheet: _write_sheet_calc_pr()";
$expected = '<sheetCalcPr fullCalcOnLoad="1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_sheet_calc_pr();

is( $got, $expected, $caption );

__END__


