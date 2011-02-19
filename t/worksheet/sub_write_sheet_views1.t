###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# reverse('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
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
# 1. Test the _write_sheet_views() method.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0" /></sheetViews>';

$worksheet = _new_worksheet(\$got);

$worksheet->select();
$worksheet->_write_sheet_views();

is( $got, $expected, $caption );


__END__
