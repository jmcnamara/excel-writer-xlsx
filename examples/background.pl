#!/usr/bin/perl -w

#######################################################################
#
# An example of setting a worksheet background image with Excel::Writer::XLSX.
#
# Copyright 2000-2024, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( 'background.xlsx' );
my $worksheet  = $workbook->add_worksheet();

$worksheet->set_background( 'republic.png' );

$workbook->close();

__END__
