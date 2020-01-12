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
use Excel::Writer::XLSX::Format;

use Test::More tests => 3;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;
my $format = Excel::Writer::XLSX::Format->new( 0 );


###############################################################################
#
# Test the _write_merge_cells() method. With $row, $col notation.
#
$caption  = " \tWorksheet: _write_merge_cells()";
$expected = '<mergeCells count="1"><mergeCell ref="B3:C3"/></mergeCells>';

$worksheet = _new_worksheet(\$got);
$worksheet->merge_range( 2, 1, 2, 2, 'Foo', $format);
$worksheet->_write_merge_cells();

is( $got, $expected, $caption );

###############################################################################
#
# Test the _write_merge_cells() method. With A1 notation.
#
$caption  = " \tWorksheet: _write_merge_cells()";
$expected = '<mergeCells count="1"><mergeCell ref="B3:C3"/></mergeCells>';

$worksheet = _new_worksheet(\$got);
$worksheet->merge_range( 'B3:C3', 'Foo', $format);
$worksheet->_write_merge_cells();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_merge_cells() method. With more than one range.
#
$caption  = " \tWorksheet: _write_merge_cells()";
$expected = '<mergeCells count="2"><mergeCell ref="B3:C3"/><mergeCell ref="A2:D2"/></mergeCells>';

$worksheet = _new_worksheet(\$got);
$worksheet->merge_range( 'B3:C3', 'Foo', $format);
$worksheet->merge_range( 'A2:D2', 'Foo', $format);
$worksheet->_write_merge_cells();

is( $got, $expected, $caption );

__END__


