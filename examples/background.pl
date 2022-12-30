#!/usr/bin/perl -w

#######################################################################
#
# An example of setting a worksheet background image with Excel::Writer::XLSX.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( 'background.xlsx' );
my $worksheet  = $workbook->add_worksheet();

$worksheet->set_background( 'republic.png' );

$workbook->close();

__END__
