#!/usr/bin/perl -w

###############################################################################
#
# Example of formatting using the Excel::Writer::XLSX module
#
# This is a simple example of how to use functions that reference cells in
# other worksheets within the same workbook.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

# Create a new workbook and add a worksheet
my $workbook   = Excel::Writer::XLSX->new( 'stats_ext.xlsx' );
my $worksheet1 = $workbook->add_worksheet( 'Test results' );
my $worksheet2 = $workbook->add_worksheet( 'Data' );

# Set the column width for columns 1
$worksheet1->set_column( 'A:A', 20 );


# Create a format for the headings
my $heading = $workbook->add_format();
$heading->set_bold();

# Create a numerical format
my $numformat = $workbook->add_format();
$numformat->set_num_format( '0.00' );


# Write some statistical functions
$worksheet1->write( 'A1', 'Count', $heading );
$worksheet1->write( 'B1', '=COUNT(Data!B2:B9)' );

$worksheet1->write( 'A2', 'Sum', $heading );
$worksheet1->write( 'B2', '=SUM(Data!B2:B9)' );

$worksheet1->write( 'A3', 'Average', $heading );
$worksheet1->write( 'B3', '=AVERAGE(Data!B2:B9)' );

$worksheet1->write( 'A4', 'Min', $heading );
$worksheet1->write( 'B4', '=MIN(Data!B2:B9)' );

$worksheet1->write( 'A5', 'Max', $heading );
$worksheet1->write( 'B5', '=MAX(Data!B2:B9)' );

$worksheet1->write( 'A6', 'Standard Deviation', $heading );
$worksheet1->write( 'B6', '=STDEV(Data!B2:B9)' );

$worksheet1->write( 'A7', 'Kurtosis', $heading );
$worksheet1->write( 'B7', '=KURT(Data!B2:B9)' );


# Write the sample data
$worksheet2->write( 'A1', 'Sample', $heading );
$worksheet2->write( 'A2', 1 );
$worksheet2->write( 'A3', 2 );
$worksheet2->write( 'A4', 3 );
$worksheet2->write( 'A5', 4 );
$worksheet2->write( 'A6', 5 );
$worksheet2->write( 'A7', 6 );
$worksheet2->write( 'A8', 7 );
$worksheet2->write( 'A9', 8 );

$worksheet2->write( 'B1', 'Length', $heading );
$worksheet2->write( 'B2', 25.4,     $numformat );
$worksheet2->write( 'B3', 25.4,     $numformat );
$worksheet2->write( 'B4', 24.8,     $numformat );
$worksheet2->write( 'B5', 25.0,     $numformat );
$worksheet2->write( 'B6', 25.3,     $numformat );
$worksheet2->write( 'B7', 24.9,     $numformat );
$worksheet2->write( 'B8', 25.2,     $numformat );
$worksheet2->write( 'B9', 24.8,     $numformat );

$workbook->close();

__END__
