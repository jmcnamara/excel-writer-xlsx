#!/usr/bin/perl

###############################################################################
#
# Simple example of merging cells using the Excel::Writer::XLSX module.
#
# This example merges three cells using the "Centre Across Selection"
# alignment which was the Excel 5 method of achieving a merge. For a more
# modern approach use the merge_range() worksheet method instead.
# See the merge3.pl - merge6.pl programs.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use warnings;
use Excel::Writer::XLSX;

# Create a new workbook and add a worksheet
my $workbook  = Excel::Writer::XLSX->new( 'merge1.xlsx' );
my $worksheet = $workbook->add_worksheet();


# Increase the cell size of the merged cells to highlight the formatting.
$worksheet->set_column( 'B:D', 20 );
$worksheet->set_row( 2, 30 );


# Create a merge format
my $format = $workbook->add_format( center_across => 1 );


# Only one cell should contain text, the others should be blank.
$worksheet->write( 2, 1, "Center across selection", $format );
$worksheet->write_blank( 2, 2, $format );
$worksheet->write_blank( 2, 3, $format );

$workbook->close();

__END__
