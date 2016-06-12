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

use Test::More tests => 8;


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
$caption  = " \tWorkbook: _write_workbook_view() 1";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>';

$workbook = _new_workbook(\$got);
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_workbook_view() method. Second tab selected.
#
$caption  = " \tWorkbook: _write_workbook_view() 2";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" activeTab="1"/>';

$workbook = _new_workbook(\$got);
$workbook->{_activesheet} = 1;
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_workbook_view() method. Second tab selected. First sheet set.
#
$caption  = " \tWorkbook: _write_workbook_view() 3";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" firstSheet="2" activeTab="1"/>';

$workbook = _new_workbook(\$got);
$workbook->{_firstsheet} = 1;
$workbook->{_activesheet} = 1;
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_workbook_view() method with set_size().
#
$caption  = " \tWorkbook: _write_workbook_view() 4";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>';

$workbook = _new_workbook(\$got);
$workbook->set_size();
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_workbook_view() method with set_size().
#
$caption  = " \tWorkbook: _write_workbook_view() 5";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>';

$workbook = _new_workbook(\$got);
$workbook->set_size(0, 0);
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_workbook_view() method with set_size().
#
$caption  = " \tWorkbook: _write_workbook_view() 6";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>';

$workbook = _new_workbook(\$got);
$workbook->set_size(1073, 644);
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_workbook_view() method with set_size().
#
$caption  = " \tWorkbook: _write_workbook_view() 7";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="1845" windowHeight="1050"/>';

$workbook = _new_workbook(\$got);
$workbook->set_size(123, 70);
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_workbook_view() method with set_size().
#
$caption  = " \tWorkbook: _write_workbook_view() 8";
$expected = '<workbookView xWindow="240" yWindow="15" windowWidth="10785" windowHeight="7350"/>';

$workbook = _new_workbook(\$got);
$workbook->set_size(719, 490);
$workbook->_write_workbook_view();

is( $got, $expected, $caption );


__END__
