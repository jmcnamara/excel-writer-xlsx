#!/usr/bin/perl -w

#######################################################################
#
# Demo of some of the features of Excel::XLSX::Writer.
# Used to create the project screenshot for Freshmeat.
#
#
# reverse('©'), October 2001, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::XLSX::Writer;

my $workbook   = Excel::XLSX::Writer->new("demo.xls");

die "Couldn't create new Excel file: $!.\n" unless defined $workbook;

my $worksheet  = $workbook->add_worksheet('Demo');
my $worksheet2 = $workbook->add_worksheet('Another sheet');
my $worksheet3 = $workbook->add_worksheet('And another');

my $bold       = $workbook->add_format(bold => 1);


#######################################################################
#
# Write a general heading
#
$worksheet->set_column('A:A', 48, $bold);
$worksheet->set_column('B:B', 20       );
$worksheet->set_row   (0,     40       );

my $heading  = $workbook->add_format(
                                        bold    => 1,
                                        color   => 'blue',
                                        size    => 16,
                                        merge   => 1,
                                        align  => 'vcenter',
                                        );

my @headings = ('Features of Excel::XLSX::Writer', '');
$worksheet->write_row('A1', \@headings, $heading);


#######################################################################
#
# Some text examples
#
my $text_format  = $workbook->add_format(
                                            bold    => 1,
                                            italic  => 1,
                                            color   => 'red',
                                            size    => 18,
                                            font    =>'Lucida Calligraphy'
                                        );

$worksheet->write('A2', "Text");
$worksheet->write('B2', "Hello Excel");
$worksheet->write('A3', "Formatted text");
$worksheet->write('B3', "Hello Excel", $text_format);

#######################################################################
#
# Some numeric examples
#
my $num1_format  = $workbook->add_format(num_format => '$#,##0.00');
my $num2_format  = $workbook->add_format(num_format => ' d mmmm yyy');


$worksheet->write('A4', "Numbers");
$worksheet->write('B4', 1234.56);
$worksheet->write('A5', "Formatted numbers");
$worksheet->write('B5', 1234.56, $num1_format);
$worksheet->write('A6', "Formatted numbers");
$worksheet->write('B6', 37257, $num2_format);


#######################################################################
#
# Formulae
#
$worksheet->set_selection('B7');
$worksheet->write('A7', 'Formulas and functions, "=SIN(PI()/4)"');
$worksheet->write('B7', '=SIN(PI()/4)');


#######################################################################
#
# Hyperlinks
#
my $url_format  = $workbook->add_format(
                                            underline => 1,
                                            color     => 'blue',
                                        );

$worksheet->write('A8', "Hyperlinks");
$worksheet->write('B8',  'http://www.perl.com/', $url_format);


#######################################################################
#
# Images
#
$worksheet->write('A9', "Images");
$worksheet->insert_bitmap('B9', 'republic.bmp', 16, 8);


#######################################################################
#
# Misc
#
$worksheet->write('A17', "Page/printer setup");
$worksheet->write('A18', "Multiple worksheets");


