###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_workbook';
use strict;
use warnings;

use Test::More tests => 3;


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
# Test the _write_workbook_view() method.
#
$caption  = " \tWorkbook: _write_workbook_view()";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" />';

$workbook = _new_workbook(\$got);
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_workbook_view() method. Second tab selected.
#
$caption  = " \tWorkbook: _write_workbook_view()";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" activeTab="1" />';

$workbook = _new_workbook(\$got);
$workbook->{_activesheet} = 1;
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_workbook_view() method. Second tab selected. First sheet set.
#
$caption  = " \tWorkbook: _write_workbook_view()";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" firstSheet="1" activeTab="1" />';

$workbook = _new_workbook(\$got);
$workbook->{_firstsheet} = 1;
$workbook->{_activesheet} = 1;
$workbook->_write_workbook_view();

is( $got, $expected, $caption );

__END__


