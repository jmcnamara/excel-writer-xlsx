#######################################################################
#
# An example of embedding images into a worksheet cells using the the
# Excel::Writer::XLSX module. This is equivalent to Excel's "Place in cell"
# image insert.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use Excel::Writer::XLSX;

# Create a new workbook called simple.xls and add a worksheet
my $workbook  = Excel::Writer::XLSX->new( 'embedded_images.xlsx' );
my $worksheet = $workbook->add_worksheet();


# Widen the first column to make the caption clearer.
$worksheet->set_column( 0, 0, 30 );
$worksheet->write( 0, 0, "Embed images that scale to cell size" );

# Embed an images in cells of different widths/heights.
$worksheet->set_column( 1, 1, 14 );

$worksheet->set_row( 1, 60 );
$worksheet->embed_image( 1, 1, "republic.png" );

$worksheet->set_row( 3, 120 );
$worksheet->embed_image( 3, 1, "republic.png" );

$workbook->close();
