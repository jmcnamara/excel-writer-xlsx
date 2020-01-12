###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
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
# Test the _write_col_breaks() method.
#
$caption  = " \tWorksheet: _write_col_breaks()";
$expected = '<colBreaks count="1" manualBreakCount="1"><brk id="1" max="1048575" man="1"/></colBreaks>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_vbreaks} = [1];
$worksheet->_write_col_breaks();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_col_breaks() method.
#
$caption  = " \tWorksheet: _write_col_breaks()";
$expected = '<colBreaks count="3" manualBreakCount="3"><brk id="1" max="1048575" man="1"/><brk id="3" max="1048575" man="1"/><brk id="8" max="1048575" man="1"/></colBreaks>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_vbreaks} = [8, 3, 1, 0];
$worksheet->_write_col_breaks();

is( $got, $expected, $caption );

__END__


