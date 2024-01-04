#!/usr/bin/perl -w

#######################################################################
#
# An example of adding a worksheet watermark image using the Excel::Writer::XLSX
# module. This is based on the method of putting an image in the worksheet
# header as suggested in the Microsoft documentation:
# https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'watermark.xlsx' );
my $worksheet = $workbook->add_worksheet();

# Set a worksheet header with the watermark image.
$worksheet->set_header( '&C&C&[Picture]', undef, { image_center => 'watermark.png' } );

$workbook->close();

__END__
