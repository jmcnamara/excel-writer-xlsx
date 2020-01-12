#!/usr/bin/perl -w

###############################################################################
#
# Example of how to add a user defined data handler to the
# Excel::Writer::XLSX write() method.
#
# The following example shows how to add a handler for a 7 digit ID number.
#
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;


my $workbook  = Excel::Writer::XLSX->new( 'write_handler1.xlsx' );
my $worksheet = $workbook->add_worksheet();


###############################################################################
#
# Add a handler for 7 digit id numbers. This is useful when you want a string
# such as 0000001 written as a string instead of a number and thus preserve
# the leading zeroes.
#
# Note: you can get the same effect using the keep_leading_zeros() method but
# this serves as a simple example.
#
$worksheet->add_write_handler( qr[^\d{7}$], \&write_my_id );


###############################################################################
#
# The following function processes the data when a match is found.
#
sub write_my_id {

    my $worksheet = shift;

    return $worksheet->write_string( @_ );
}


# This format maintains the cell as text even if it is edited.
my $id_format = $workbook->add_format( num_format => '@' );


# Write some numbers in the user defined format
$worksheet->write( 'A1', '0000000', $id_format );
$worksheet->write( 'A2', '0000001', $id_format );
$worksheet->write( 'A3', '0004000', $id_format );
$worksheet->write( 'A4', '1234567', $id_format );

# Write some numbers that don't match the defined format
$worksheet->write( 'A6', '000000', $id_format );
$worksheet->write( 'A7', '000001', $id_format );
$worksheet->write( 'A8', '004000', $id_format );
$worksheet->write( 'A9', '123456', $id_format );

$workbook->close();

__END__

