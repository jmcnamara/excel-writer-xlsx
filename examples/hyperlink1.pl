#!/usr/bin/perl

###############################################################################
#
# Example of how to use the Excel::Writer::XLSX module to write hyperlinks
#
# See also hyperlink2.pl for worksheet URL examples.
#
# reverse('©'), May 2004, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

# Create a new workbook and add a worksheet
my $workbook = Excel::Writer::XLSX->new( 'hyperlink.xlsx' );


my $worksheet = $workbook->add_worksheet( 'Hyperlinks' );

# Format the first column
$worksheet->set_column( 'A:A', 30 );
$worksheet->set_selection( 'B1' );


# Add the standard url link format.
my $url_format = $workbook->add_format(
    color     => 'blue',
    underline => 1,
);

# Add a sample format.
my $red_format = $workbook->add_format(
    color     => 'red',
    bold      => 1,
    underline => 1,
    size      => 12,
);

# Add an alternate description string to the URL.
my $str = 'Perl home.';

# Add a "tool tip" to the URL.
my $tip = 'Get the latest Perl news here.';


# Write some hyperlinks
$worksheet->write( 'A1', 'http://www.perl.com/', $url_format );
$worksheet->write( 'A3', 'http://www.perl.com/', $url_format, $str );
$worksheet->write( 'A5', 'http://www.perl.com/', $url_format, $str, $tip );
$worksheet->write( 'A7', 'http://www.perl.com/', $red_format );
$worksheet->write( 'A9', 'mailto:jmcnamara@cpan.org', $url_format, 'Mail me' );

# Write a URL that isn't a hyperlink
$worksheet->write_string( 'A11', 'http://www.perl.com/' );

