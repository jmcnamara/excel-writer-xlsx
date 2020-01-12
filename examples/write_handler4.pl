#!/usr/bin/perl -w

###############################################################################
#
# Example of how to add a user defined data handler to the
# Excel::Writer::XLSX write() method.
#
# The following example shows how to add a handler for dates in a specific
# format.
#
# This is a more rigorous version of write_handler3.pl.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;


my $workbook    = Excel::Writer::XLSX->new( 'write_handler4.xlsx' );
my $worksheet   = $workbook->add_worksheet();
my $date_format = $workbook->add_format( num_format => 'dd/mm/yy' );


###############################################################################
#
# Add a handler to match dates in the following formats: d/m/yy, d/m/yyyy
#
# The day and month can be single or double digits and the year can be  2 or 4
# digits.
#
$worksheet->add_write_handler( qr[^\d{1,2}/\d{1,2}/\d{2,4}$], \&write_my_date );


###############################################################################
#
# The following function processes the data when a match is found.
#
sub write_my_date {

    my $worksheet = shift;
    my @args      = @_;

    my $token = $args[2];

    if ( $token =~ qr[^(\d{1,2})/(\d{1,2})/(\d{2,4})$] ) {

        my $day  = $1;
        my $mon  = $2;
        my $year = $3;

        # Use a window for 2 digit dates. This will keep some ragged Perl
        # programmer employed in thirty years time. :-)
        if ( length $year == 2 ) {
            if ( $year < 50 ) {
                $year += 2000;
            }
            else {
                $year += 1900;
            }
        }

        my $date = sprintf "%4d-%02d-%02dT", $year, $mon, $day;

        # Convert the ISO ISO8601 style string to an Excel date
        $date = $worksheet->convert_date_time( $date );

        if ( defined $date ) {

            # Date was valid
            $args[2] = $date;
            return $worksheet->write_number( @args );
        }
        else {

            # Not a valid date therefore write as a string
            return $worksheet->write_string( @args );
        }
    }
    else {

        # Shouldn't happen if the same match is used in the re and sub.
        return undef;
    }
}


# Write some dates in the user defined format
$worksheet->write( 'A1', '22/12/2004', $date_format );
$worksheet->write( 'A2', '22/12/04',   $date_format );
$worksheet->write( 'A3', '2/12/04',    $date_format );
$worksheet->write( 'A4', '2/5/04',     $date_format );
$worksheet->write( 'A5', '2/5/95',     $date_format );
$worksheet->write( 'A6', '2/5/1995',   $date_format );

# Some erroneous dates
$worksheet->write( 'A8', '2/5/1895',  $date_format ); # Date out of Excel range
$worksheet->write( 'A9', '29/2/2003', $date_format ); # Invalid leap day
$worksheet->write( 'A10', '50/50/50', $date_format ); # Matches but isn't a date

$workbook->close();

__END__

