#!/usr/bin/perl

###############################################################################
#
# This example demonstrates writing cell comments.
#
# A cell comment is indicated in Excel by a small red triangle in the upper
# right-hand corner of the cell.
#
# For more advanced comment options see comments2.pl.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'comments1.xlsx' );
my $worksheet = $workbook->add_worksheet();


$worksheet->write( 'A1', 'Hello' );
$worksheet->write_comment( 'A1', 'This is a comment' );

$workbook->close();

__END__
