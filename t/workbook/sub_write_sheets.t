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
# Test the _write_sheets() method.
#
$caption  = " \tWorkbook: _write_sheets()";
$expected = '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>';

$workbook = _new_workbook(\$got);
my $worksheet = $workbook->add_worksheet();

$workbook->_write_sheets();

is( $got, $expected, $caption );

__END__


