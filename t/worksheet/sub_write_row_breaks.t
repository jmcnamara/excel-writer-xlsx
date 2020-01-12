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
# Test the _write_row_breaks() method.
#
$caption  = " \tWorksheet: _write_row_breaks()";
$expected = '<rowBreaks count="1" manualBreakCount="1"><brk id="1" max="16383" man="1"/></rowBreaks>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_hbreaks} = [1];
$worksheet->_write_row_breaks();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_row_breaks() method.
#
$caption  = " \tWorksheet: _write_row_breaks()";
$expected = '<rowBreaks count="3" manualBreakCount="3"><brk id="3" max="16383" man="1"/><brk id="7" max="16383" man="1"/><brk id="15" max="16383" man="1"/></rowBreaks>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_hbreaks} = [15, 7, 3, 0];
$worksheet->_write_row_breaks();

is( $got, $expected, $caption );

__END__


