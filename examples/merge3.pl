#!/usr/bin/perl

###############################################################################
#
# Example of how to use Excel::Writer::XLSX to write a hyperlink in a
# merged cell. There are two options write_url_range() with a standard merge
# format or merge_range().
#
# reverse('©'), September 2002, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

# Create a new workbook and add a worksheet
my $workbook  = Excel::Writer::XLSX->new( 'merge3.xlsx' );
my $worksheet = $workbook->add_worksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_row( $_, 30 ) for ( 1, 3, 6, 7 );
$worksheet->set_column( 'B:D', 20 );


###############################################################################
#
# Example 1: Merge cells containing a hyperlink using write_url_range()
# and the standard Excel 5+ merge property.
#
my $format1 = $workbook->add_format(
    center_across => 1,
    border        => 1,
    underline     => 1,
    color         => 'blue',
);

# Write the cells to be merged
$worksheet->write_url_range( 'B2:D2', 'http://www.perl.com', $format1 );
$worksheet->write_blank( 'C2', $format1 );
$worksheet->write_blank( 'D2', $format1 );


###############################################################################
#
# Example 2: Merge cells containing a hyperlink using merge_range().
#
my $format2 = $workbook->add_format(
    border    => 1,
    underline => 1,
    color     => 'blue',
    align     => 'center',
    valign    => 'vcenter',
);

# Merge 3 cells
$worksheet->merge_range( 'B4:D4', 'http://www.perl.com', $format2 );


# Merge 3 cells over two rows
$worksheet->merge_range( 'B7:D8', 'http://www.perl.com', $format2 );


