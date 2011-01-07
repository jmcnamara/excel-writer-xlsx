###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# reverse('©'), January 2011, John McNamara, jmcnamara@cpan.org
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
# Test the _write_defined_names() method.
#
$caption  = " \tWorkbook: _write_defined_names()";
$expected = '<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName></definedNames>';

$workbook = _new_workbook(\$got);
$workbook->{_defined_names} = [ [ '_xlnm.Print_Titles', 0, 'Sheet1!$1:$1' ] ];

$workbook->_write_defined_names();

is( $got, $expected, $caption );

__END__


