###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions '_new_workbook';
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
my $workbook;


###############################################################################
#
# Test the _write_workbook() method.
#
$caption  = " \tWorkbook: _write_workbook()";
$expected = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

$workbook = _new_workbook(\$got);

$workbook->_write_workbook();

is( $got, $expected, $caption );

__END__


