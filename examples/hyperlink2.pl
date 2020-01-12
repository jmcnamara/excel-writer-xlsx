#!/usr/bin/perl

###############################################################################
#
# Example of how to use the Excel::Writer::XLSX module to write internal and
# external hyperlinks.
#
# If you wish to run this program and follow the hyperlinks you should create
# the following directory structure:
#
# C:\ -- Temp --+-- Europe
#               |
#               \-- Asia
#
#
# See also hyperlink1.pl for web URL examples.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#


use strict;
use warnings;
use Excel::Writer::XLSX;

# Create three workbooks:
#   C:\Temp\Europe\Ireland.xlsx
#   C:\Temp\Europe\Italy.xlsx
#   C:\Temp\Asia\China.xlsx
#

my $ireland = Excel::Writer::XLSX->new( 'C:\Temp\Europe\Ireland.xlsx' );

my $ire_links      = $ireland->add_worksheet( 'Links' );
my $ire_sales      = $ireland->add_worksheet( 'Sales' );
my $ire_data       = $ireland->add_worksheet( 'Product Data' );
my $ire_url_format = $ireland->get_default_url_format();


my $italy = Excel::Writer::XLSX->new( 'C:\Temp\Europe\Italy.xlsx' );

my $ita_links      = $italy->add_worksheet( 'Links' );
my $ita_sales      = $italy->add_worksheet( 'Sales' );
my $ita_data       = $italy->add_worksheet( 'Product Data' );
my $ita_url_format = $italy->get_default_url_format();


my $china = Excel::Writer::XLSX->new( 'C:\Temp\Asia\China.xlsx' );

my $cha_links      = $china->add_worksheet( 'Links' );
my $cha_sales      = $china->add_worksheet( 'Sales' );
my $cha_data       = $china->add_worksheet( 'Product Data' );
my $cha_url_format = $china->get_default_url_format();


# Add an alternative format
my $format = $ireland->add_format( color => 'green', bold => 1 );
$ire_links->set_column( 'A:B', 25 );


###############################################################################
#
# Examples of internal links
#
$ire_links->write( 'A1', 'Internal links', $format );

# Internal link
$ire_links->write_url( 'A2', 'internal:Sales!A2', $ire_url_format );

# Internal link to a range
$ire_links->write_url( 'A3', 'internal:Sales!A3:D3', $ire_url_format );

# Internal link with an alternative string
$ire_links->write_url( 'A4', 'internal:Sales!A4', $ire_url_format, 'Link' );

# Internal link with an alternative format
$ire_links->write_url( 'A5', 'internal:Sales!A5', $format );

# Internal link with an alternative string and format
$ire_links->write_url( 'A6', 'internal:Sales!A6', $ire_url_format, 'Link' );

# Internal link (spaces in worksheet name)
$ire_links->write_url( 'A7', q{internal:'Product Data'!A7}, $ire_url_format );


###############################################################################
#
# Examples of external links
#
$ire_links->write( 'B1', 'External links', $format );

# External link to a local file
$ire_links->write_url( 'B2', 'external:Italy.xlsx', $ire_url_format );

# External link to a local file with worksheet
$ire_links->write_url( 'B3', 'external:Italy.xlsx#Sales!B3', $ire_url_format );

# External link to a local file with worksheet and alternative string
$ire_links->write_url( 'B4', 'external:Italy.xlsx#Sales!B4', $ire_url_format, 'Link' );

# External link to a local file with worksheet and format
$ire_links->write_url( 'B5', 'external:Italy.xlsx#Sales!B5', $format );

# External link to a remote file, absolute path
$ire_links->write_url( 'B6', 'external:C:/Temp/Asia/China.xlsx', $ire_url_format );

# External link to a remote file, relative path
$ire_links->write_url( 'B7', 'external:../Asia/China.xlsx', $ire_url_format );

# External link to a remote file with worksheet
$ire_links->write_url( 'B8', 'external:C:/Temp/Asia/China.xlsx#Sales!B8', $ire_url_format );

# External link to a remote file with worksheet (with spaces in the name)
$ire_links->write_url( 'B9', q{external:C:/Temp/Asia/China.xlsx#'Product Data'!B9}, $ire_url_format );


###############################################################################
#
# Some utility links to return to the main sheet
#
$ire_sales->write_url( 'A2', 'internal:Links!A2', $ire_url_format, 'Back' );
$ire_sales->write_url( 'A3', 'internal:Links!A3', $ire_url_format, 'Back' );
$ire_sales->write_url( 'A4', 'internal:Links!A4', $ire_url_format, 'Back' );
$ire_sales->write_url( 'A5', 'internal:Links!A5', $ire_url_format, 'Back' );
$ire_sales->write_url( 'A6', 'internal:Links!A6', $ire_url_format, 'Back' );
$ire_data->write_url ( 'A7', 'internal:Links!A7', $ire_url_format, 'Back' );

$ita_links->write_url( 'A1', 'external:Ireland.xlsx#Links!B2', $ita_url_format, 'Back' );
$ita_sales->write_url( 'B3', 'external:Ireland.xlsx#Links!B3', $ita_url_format, 'Back' );
$ita_sales->write_url( 'B4', 'external:Ireland.xlsx#Links!B4', $ita_url_format, 'Back' );
$ita_sales->write_url( 'B5', 'external:Ireland.xlsx#Links!B5', $ita_url_format, 'Back' );
$cha_links->write_url( 'A1', 'external:C:/Temp/Europe/Ireland.xlsx#Links!B6', $cha_url_format, 'Back' );
$cha_sales->write_url( 'B8', 'external:C:/Temp/Europe/Ireland.xlsx#Links!B8', $cha_url_format, 'Back' );
$cha_data->write_url ( 'B9', 'external:C:/Temp/Europe/Ireland.xlsx#Links!B9', $cha_url_format, 'Back' );

$ireland->close();
$italy->close();
$china->close();

__END__
