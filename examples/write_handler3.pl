#!/usr/bin/perl -w

###############################################################################
#
# Example of how to add a user defined data handler to the
# Excel::Writer::XLSX write() method.
#
# The following example shows how to add a handler for dates in a specific
# format.
#
# See write_handler4.pl for a more rigorous example with error handling.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;


my $workbook    = Excel::Writer::XLSX->new( 'write_handler3.xlsx' );
my $worksheet   = $workbook->add_worksheet();
my $date_format = $workbook->add_format( num_format => 'dd/mm/yy' );


###############################################################################
#
# Add a handler to match dates in the following format: d/m/yyyy
#
# The day and month can be single or double digits.
#
$worksheet->add_write_handler( qr[^\d{1,2}/\d{1,2}/\d{4}$], \&write_my_date );


###############################################################################
#
# The following function processes the data when a match is found.
# See write_handler4.pl for a more rigorous example with error handling.
#
sub write_my_date {

    my $worksheet = shift;
    my @args      = @_;

    my $token = $args[2];
    $token =~ qr[^(\d{1,2})/(\d{1,2})/(\d{4})$];

    # Change to the date format required by write_date_time().
    my $date = sprintf "%4d-%02d-%02dT", $3, $2, $1;

    $args[2] = $date;

    return $worksheet->write_date_time( @args );
}


# Write some dates in the user defined format
$worksheet->write( 'A1', '22/12/2004', $date_format );
$worksheet->write( 'A2', '1/1/1995',   $date_format );
$worksheet->write( 'A3', '01/01/1995', $date_format );

$workbook->close();

__END__

