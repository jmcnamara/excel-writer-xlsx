#!/usr/bin/perl

#######################################################################
#
# Example of how to hide a worksheet with Excel::Writer::XLSX.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( 'hide_sheet.xlsx' );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet();

$worksheet1->set_column( 'A:A', 30 );
$worksheet2->set_column( 'A:A', 30 );
$worksheet3->set_column( 'A:A', 30 );

# Sheet2 won't be visible until it is unhidden in Excel.
$worksheet2->hide();

$worksheet1->write( 0, 0, 'Sheet2 is hidden' );
$worksheet2->write( 0, 0, "Now it's my turn to find you." );
$worksheet3->write( 0, 0, 'Sheet2 is hidden' );

$workbook->close();

__END__
