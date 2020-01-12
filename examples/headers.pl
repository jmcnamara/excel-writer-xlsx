#!/usr/bin/perl

######################################################################
#
# This program shows several examples of how to set up headers and
# footers with Excel::Writer::XLSX.
#
# The control characters used in the header/footer strings are:
#
#     Control             Category            Description
#     =======             ========            ===========
#     &L                  Justification       Left
#     &C                                      Center
#     &R                                      Right
#
#     &P                  Information         Page number
#     &N                                      Total number of pages
#     &D                                      Date
#     &T                                      Time
#     &F                                      File name
#     &A                                      Worksheet name
#
#     &fontsize           Font                Font size
#     &"font,style"                           Font name and style
#     &U                                      Single underline
#     &E                                      Double underline
#     &S                                      Strikethrough
#     &X                                      Superscript
#     &Y                                      Subscript
#
#     &[Picture]          Images              Image placeholder
#     &G                                      Same as &[Picture]
#
#     &&                  Miscellaneous       Literal ampersand &
#
# See the main Excel::Writer::XLSX documentation for more information.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#


use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook = Excel::Writer::XLSX->new( 'headers.xlsx' );
my $preview  = 'Select Print Preview to see the header and footer';


######################################################################
#
# A simple example to start
#
my $worksheet1 = $workbook->add_worksheet( 'Simple' );
my $header1    = '&CHere is some centred text.';
my $footer1    = '&LHere is some left aligned text.';

$worksheet1->set_header( $header1 );
$worksheet1->set_footer( $footer1 );

$worksheet1->set_column( 'A:A', 50 );
$worksheet1->write( 'A1', $preview );


######################################################################
#
# A simple example to start
#
my $worksheet2 = $workbook->add_worksheet( 'Image' );
my $header2    = '&L&[Picture]';

# Adjust the page top margin to allow space for the header image.
$worksheet2->set_margin_top(1.75);

$worksheet2->set_header( $header2, 0.3, {image_left => 'republic.png'});

$worksheet2->set_column( 'A:A', 50 );
$worksheet2->write( 'A1', $preview );


######################################################################
#
# This is an example of some of the header/footer variables.
#
my $worksheet3 = $workbook->add_worksheet( 'Variables' );
my $header3    = '&LPage &P of &N' . '&CFilename: &F' . '&RSheetname: &A';
my $footer3    = '&LCurrent date: &D' . '&RCurrent time: &T';

$worksheet3->set_header( $header3 );
$worksheet3->set_footer( $footer3 );

$worksheet3->set_column( 'A:A', 50 );
$worksheet3->write( 'A1',  $preview );
$worksheet3->write( 'A21', 'Next sheet' );
$worksheet3->set_h_pagebreaks( 20 );


######################################################################
#
# This example shows how to use more than one font
#
my $worksheet4 = $workbook->add_worksheet( 'Mixed fonts' );
my $header4    = q(&C&"Courier New,Bold"Hello &"Arial,Italic"World);
my $footer4    = q(&C&"Symbol"e&"Arial" = mc&X2);

$worksheet4->set_header( $header4 );
$worksheet4->set_footer( $footer4 );

$worksheet4->set_column( 'A:A', 50 );
$worksheet4->write( 'A1', $preview );


######################################################################
#
# Example of line wrapping
#
my $worksheet5 = $workbook->add_worksheet( 'Word wrap' );
my $header5    = "&CHeading 1\nHeading 2";

$worksheet5->set_header( $header5 );

$worksheet5->set_column( 'A:A', 50 );
$worksheet5->write( 'A1', $preview );


######################################################################
#
# Example of inserting a literal ampersand &
#
my $worksheet6 = $workbook->add_worksheet( 'Ampersand' );
my $header6    = '&CCuriouser && Curiouser - Attorneys at Law';

$worksheet6->set_header( $header6 );

$worksheet6->set_column( 'A:A', 50 );
$worksheet6->write( 'A1', $preview );

$workbook->close();

__END__
