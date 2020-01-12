#!/usr/bin/perl

###############################################################################
#
# Example of how to use the Excel::Writer::XLSX merge_cells() workbook
# method with Unicode strings.
#
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

# Create a new workbook and add a worksheet
my $workbook  = Excel::Writer::XLSX->new( 'merge6.xlsx' );
my $worksheet = $workbook->add_worksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_row( $_, 36 ) for 2 .. 9;
$worksheet->set_column( 'B:D', 25 );


# Format for the merged cells.
my $format = $workbook->add_format(
    border => 6,
    bold   => 1,
    color  => 'red',
    size   => 20,
    valign => 'vcentre',
    align  => 'left',
    indent => 1,
);


###############################################################################
#
# Write an Ascii string.
#
$worksheet->merge_range( 'B3:D4', 'ASCII: A simple string', $format );


###############################################################################
#
# Write a UTF-8 Unicode string.
#
my $smiley = chr 0x263a;
$worksheet->merge_range( 'B6:D7', "UTF-8: A Unicode smiley $smiley", $format );

$workbook->close();

__END__
