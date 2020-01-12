#!/usr/bin/perl

#######################################################################
#
# Example of how to use the Excel::Writer::XLSX module to write simple
# array formulas.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

# Create a new workbook and add a worksheet
my $workbook  = Excel::Writer::XLSX->new( 'array_formula.xlsx' );
my $worksheet = $workbook->add_worksheet();

# Write some test data.
$worksheet->write( 'B1', [ [ 500, 10 ], [ 300, 15 ] ] );
$worksheet->write( 'B5', [ [ 1, 2, 3 ], [ 20234, 21003, 10000 ] ] );

# Write an array formula that returns a single value
$worksheet->write( 'A1', '{=SUM(B1:C1*B2:C2)}' );

# Same as above but more verbose.
$worksheet->write_array_formula( 'A2:A2', '{=SUM(B1:C1*B2:C2)}' );

# Write an array formula that returns a range of values
$worksheet->write_array_formula( 'A5:A7', '{=TREND(C5:C7,B5:B7)}' );

$workbook->close();

__END__


