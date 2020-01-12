#!/usr/bin/perl

###############################################################################
#
# Example of how to hide rows and columns in Excel::Writer::XLSX. In order to
# hide rows without setting each one, (of approximately 1 million rows),
# Excel uses an optimisation to hide all rows that don't have data.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'hide_row_col.xlsx' );
my $worksheet = $workbook->add_worksheet();


# Write some data.
$worksheet->write( 'D1', 'Some hidden columns.' );
$worksheet->write( 'A8', 'Some hidden rows.' );

# Hide all rows without data.
$worksheet->set_default_row( undef, 1 );

# Set emptys row that we do want to display. All other will be hidden.
for my $row (1 .. 6) {
    $worksheet->set_row( $row, 15 );
}

# Hide a range of columns.
$worksheet->set_column( 'G:XFD', undef, undef, 1);

$workbook->close();

__END__



