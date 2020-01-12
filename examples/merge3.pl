#!/usr/bin/perl

###############################################################################
#
# Example of how to use Excel::Writer::XLSX to write a hyperlink in a
# merged cell.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

# Create a new workbook and add a worksheet
my $workbook  = Excel::Writer::XLSX->new( 'merge3.xlsx' );
my $worksheet = $workbook->add_worksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_row( $_, 30 ) for ( 3, 6, 7 );
$worksheet->set_column( 'B:D', 20 );


###############################################################################
#
# Example: Merge cells containing a hyperlink using merge_range().
#
my $format = $workbook->add_format(
    border    => 1,
    underline => 1,
    color     => 'blue',
    align     => 'center',
    valign    => 'vcenter',
);

# Merge 3 cells
$worksheet->merge_range( 'B4:D4', 'http://www.perl.com', $format );


# Merge 3 cells over two rows
$worksheet->merge_range( 'B7:D8', 'http://www.perl.com', $format );


$workbook->close();

__END__
