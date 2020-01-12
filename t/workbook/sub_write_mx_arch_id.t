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
# Test the _write_mx_arch_id() method.
#
$caption  = " \tWorkbook: _write_mx_arch_id()";
$expected = '<mx:ArchID Flags="2"/>';

$workbook = _new_workbook(\$got);

$workbook->_write_mx_arch_id();

is( $got, $expected, $caption );

__END__


