###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# reverse ('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
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
# Test the _write_calc_pr() method.
#
$caption  = " \tWorkbook: _write_calc_pr()";
$expected = '<calcPr calcId="124519" fullCalcOnLoad="1"/>';

$workbook = _new_workbook(\$got);

$workbook->_write_calc_pr();

is( $got, $expected, $caption );

__END__


