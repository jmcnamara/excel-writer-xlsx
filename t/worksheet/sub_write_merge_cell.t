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
# Test the _write_merge_cell() method.
#
$caption  = " \tWorksheet: _write_merge_cell()";
$expected = '<mergeCell ref="B3:C3"/>';

$worksheet = _new_worksheet(\$got);
$worksheet->_write_merge_cell( [ 2, 1, 2, 2 ] );

is( $got, $expected, $caption );

__END__


