#!/usr/bin/perl

##############################################################################
#
# Example of how to create defined names in an Excel::Writer::XLSX file.
#
# This method is used to define a user friendly name to represent a value,
# a single cell or a range of cells in a workbook.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( 'defined_name.xlsx' );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();

# Define some global/workbook names.
$workbook->define_name( 'Exchange_rate', '=0.96' );
$workbook->define_name( 'Sales',         '=Sheet1!$G$1:$H$10' );

# Define a local/worksheet name.
$workbook->define_name( 'Sheet2!Sales', '=Sheet2!$G$1:$G$10' );

# Write some text in the file and one of the defined names in a formula.
for my $worksheet ( $workbook->sheets() ) {
    $worksheet->set_column( 'A:A', 45 );
    $worksheet->write( 'A1', 'This worksheet contains some defined names.' );
    $worksheet->write( 'A2', 'See Formulas -> Name Manager above.' );
    $worksheet->write( 'A3', 'Example formula in cell B3 ->' );

    $worksheet->write( 'B3', '=Exchange_rate' );
}

$workbook->close();

__END__
