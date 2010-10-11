#!/usr/bin/perl -w

###############################################################################
#
# Example of how to use the Excel::Writer::XLSX module to send an Excel
# file to a browser in a CGI program.
#
# On Windows the hash-bang line should be something like:
#
#     #!C:\Perl\bin\perl.exe
#
# The "Content-Disposition" line will cause a prompt to be generated to save
# the file. If you want to stream the file to the browser instead, comment out
# that line as shown below.
#
# reverse('©'), March 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

# Set the filename and send the content type
my $filename = "cgitest.xlsx";

print "Content-type: application/vnd.ms-excel\n";

# The Content-Disposition will generate a prompt to save the file. If you want
# to stream the file to the browser, comment out the following line.
print "Content-Disposition: attachment; filename=$filename\n";
print "\n";

# Create a new workbook and add a worksheet. The special Perl filehandle - will
# redirect the output to STDOUT
#
my $workbook  = Excel::Writer::XLSX->new( "-" );
my $worksheet = $workbook->add_worksheet();


# Set the column width for column 1
$worksheet->set_column( 0, 0, 20 );


# Create a format
my $format = $workbook->add_format();
$format->set_bold();
$format->set_size( 15 );
$format->set_color( 'blue' );


# Write to the workbook
$worksheet->write( 0, 0, "Hi Excel!", $format );

__END__
