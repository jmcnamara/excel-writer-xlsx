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
# Test the _write_ext_lst() method.
#
$caption  = " \tWorkbook: _write_ext_lst()";
$expected = '<extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:ArchID Flags="2"/></ext></extLst>';

$workbook = _new_workbook(\$got);

$workbook->_write_ext_lst();

is( $got, $expected, $caption );

__END__


