###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# reverse('�'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 2;

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
# Test the _write_sheet_view() method. Tab not selected.
#
$caption  = " \tWorksheet: _write_sheet_view()";
$expected = '<sheetView workbookViewId="0" />';

$worksheet = _new_worksheet(\$got);
$worksheet->_write_sheet_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_sheet_view() method. Tab selected.
#
$caption  = " \tWorksheet: _write_sheet_view()";
$expected = '<sheetView tabSelected="1" workbookViewId="0" />';

$worksheet = _new_worksheet(\$got);
$worksheet->select();
$worksheet->_write_sheet_view();

is( $got, $expected, $caption );

__END__


