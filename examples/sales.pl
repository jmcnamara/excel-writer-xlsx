#!/usr/bin/perl -w

###############################################################################
#
# Example of a sales worksheet to demonstrate several different features.
# Also uses functions from the L<Excel::Writer::XLSX::Utility> module.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;

# Create a new workbook and add a worksheet
my $workbook  = Excel::Writer::XLSX->new( 'sales.xlsx' );
my $worksheet = $workbook->add_worksheet( 'May Sales' );


# Set up some formats
my %heading = (
    bold     => 1,
    pattern  => 1,
    fg_color => '#C3FFC0',
    border   => 1,
    align    => 'center',
);

my %total = (
    bold       => 1,
    top        => 1,
    num_format => '$#,##0.00'
);

my $heading      = $workbook->add_format( %heading );
my $total_format = $workbook->add_format( %total );
my $price_format = $workbook->add_format( num_format => '$#,##0.00' );
my $date_format  = $workbook->add_format( num_format => 'mmm d yyy' );


# Write the main headings
$worksheet->freeze_panes( 1 );    # Freeze the first row
$worksheet->write( 'A1', 'Item',     $heading );
$worksheet->write( 'B1', 'Quantity', $heading );
$worksheet->write( 'C1', 'Price',    $heading );
$worksheet->write( 'D1', 'Total',    $heading );
$worksheet->write( 'E1', 'Date',     $heading );

# Set the column widths
$worksheet->set_column( 'A:A', 25 );
$worksheet->set_column( 'B:B', 10 );
$worksheet->set_column( 'C:E', 16 );


# Extract the sales data from the __DATA__ section at the end of the file.
# In reality this information would probably come from a database
my @sales;

foreach my $line ( <DATA> ) {
    chomp $line;
    next if $line eq '';

    # Simple-minded processing of CSV data. Refer to the Text::CSV_XS
    # and Text::xSV modules for a more complete CSV handling.
    my @items = split /,/, $line;
    push @sales, \@items;
}


# Write out the items from each row
my $row = 1;
foreach my $sale ( @sales ) {

    $worksheet->write( $row, 0, @$sale[0] );
    $worksheet->write( $row, 1, @$sale[1] );
    $worksheet->write( $row, 2, @$sale[2], $price_format );

    # Create a formula like '=B2*C2'
    my $formula =
      '=' . xl_rowcol_to_cell( $row, 1 ) . "*" . xl_rowcol_to_cell( $row, 2 );

    $worksheet->write( $row, 3, $formula, $price_format );

    # Parse the date
    my $date = xl_decode_date_US( @$sale[3] );
    $worksheet->write( $row, 4, $date, $date_format );
    $row++;
}

# Create a formula to sum the totals, like '=SUM(D2:D6)'
my $total = '=SUM(D2:' . xl_rowcol_to_cell( $row - 1, 3 ) . ")";

$worksheet->write( $row, 3, $total, $total_format );

$workbook->close();

__DATA__
586 card,20,125.50,5/12/01
Flat Screen Monitor,1,1300.00,5/12/01
64 MB dimms,45,49.99,5/13/01
15 GB HD,12,300.00,5/13/01
Speakers (pair),5,15.50,5/14/01

