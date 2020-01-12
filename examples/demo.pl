#!/usr/bin/perl -w

#######################################################################
#
# A simple demo of some of the features of Excel::Writer::XLSX.
#
# This program is used to create the project screenshot for Freshmeat:
# L<http://freshmeat.net/projects/writeexcel/>
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( 'demo.xlsx' );
my $worksheet  = $workbook->add_worksheet( 'Demo' );
my $worksheet2 = $workbook->add_worksheet( 'Another sheet' );
my $worksheet3 = $workbook->add_worksheet( 'And another' );

my $bold = $workbook->add_format( bold => 1 );


#######################################################################
#
# Write a general heading
#
$worksheet->set_column( 'A:A', 36, $bold );
$worksheet->set_column( 'B:B', 20 );
$worksheet->set_row( 0, 40 );

my $heading = $workbook->add_format(
    bold  => 1,
    color => 'blue',
    size  => 16,
    merge => 1,
    align => 'vcenter',
);

my @headings = ( 'Features of Excel::Writer::XLSX', '' );
$worksheet->write_row( 'A1', \@headings, $heading );


#######################################################################
#
# Some text examples
#
my $text_format = $workbook->add_format(
    bold   => 1,
    italic => 1,
    color  => 'red',
    size   => 18,
    font   => 'Lucida Calligraphy'
);


$worksheet->write( 'A2', "Text" );
$worksheet->write( 'B2', "Hello Excel" );
$worksheet->write( 'A3', "Formatted text" );
$worksheet->write( 'B3', "Hello Excel", $text_format );
$worksheet->write( 'A4', "Unicode text" );
$worksheet->write( 'B4', "\x{0410} \x{0411} \x{0412} \x{0413} \x{0414}" );

#######################################################################
#
# Some numeric examples
#
my $num1_format = $workbook->add_format( num_format => '$#,##0.00' );
my $num2_format = $workbook->add_format( num_format => ' d mmmm yyy' );


$worksheet->write( 'A5', "Numbers" );
$worksheet->write( 'B5', 1234.56 );
$worksheet->write( 'A6', "Formatted numbers" );
$worksheet->write( 'B6', 1234.56, $num1_format );
$worksheet->write( 'A7', "Formatted numbers" );
$worksheet->write( 'B7', 37257, $num2_format );


#######################################################################
#
# Formulae
#
$worksheet->set_selection( 'B8' );
$worksheet->write( 'A8', 'Formulas and functions, "=SIN(PI()/4)"' );
$worksheet->write( 'B8', '=SIN(PI()/4)' );


#######################################################################
#
# Hyperlinks
#
$worksheet->write( 'A9', "Hyperlinks" );
$worksheet->write( 'B9', 'http://www.perl.com/' );


#######################################################################
#
# Images
#
$worksheet->write( 'A10', "Images" );
$worksheet->insert_image( 'B10', 'republic.png',
                                 { x_offset => 16, y_offset => 8 } );


#######################################################################
#
# Misc
#
$worksheet->write( 'A18', "Page/printer setup" );
$worksheet->write( 'A19', "Multiple worksheets" );

$workbook->close();

__END__
