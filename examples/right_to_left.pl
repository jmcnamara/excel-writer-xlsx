#!/usr/bin/perl

#######################################################################
#
# Example of how to change the default worksheet direction from
# left-to-right to right-to-left as required by some eastern verions
# of Excel.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( 'right_to_left.xlsx' );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();

$worksheet2->right_to_left();

$worksheet1->write( 0, 0, 'Hello' );    #  A1, B1, C1, ...
$worksheet2->write( 0, 0, 'Hello' );    # ..., C1, B1, A1

$workbook->close();

__END__
