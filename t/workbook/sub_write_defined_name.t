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
# Test the _write_defined_name() method.
#
$caption  = " \tWorkbook: _write_defined_name()";
$expected = '<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName>';

$workbook = _new_workbook(\$got);

$workbook->_write_defined_name( [ '_xlnm.Print_Titles', 0, 'Sheet1!$1:$1' ] );

is( $got, $expected, $caption );

__END__


