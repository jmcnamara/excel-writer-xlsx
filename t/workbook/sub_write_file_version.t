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
# Test the _write_file_version() method.
#
$caption  = " \tWorkbook: _write_file_version()";
$expected = '<fileVersion appName="xl" lastEdited="4" '
  . 'lowestEdited="4" rupBuild="4505"/>';

$workbook = _new_workbook(\$got);

$workbook->_write_file_version();

is( $got, $expected, $caption );

__END__


