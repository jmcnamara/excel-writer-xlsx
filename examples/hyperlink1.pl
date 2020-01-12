#!/usr/bin/perl

###############################################################################
#
# Example of how to use the Excel::Writer::XLSX module to write hyperlinks
#
# See also hyperlink2.pl for worksheet URL examples.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
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

# Add a user defined hyperlink format.
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


# Write some hyperlinks. Unspecified or undefined format paraamters will be
# replace with the defuault Excel hyperlink style.
$worksheet->write( 'A1', 'http://www.perl.com/' );
$worksheet->write( 'A3', 'http://www.perl.com/', undef, $str );
$worksheet->write( 'A5', 'http://www.perl.com/', undef, $str, $tip );
$worksheet->write( 'A7', 'http://www.perl.com/', $red_format );
$worksheet->write( 'A9', 'mailto:jmcnamara@cpan.org', undef, 'Mail me' );

# Write a URL that isn't a hyperlink
$worksheet->write_string( 'A11', 'http://www.perl.com/' );

$workbook->close();

__END__
