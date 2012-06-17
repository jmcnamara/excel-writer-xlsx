#!/usr/bin/perl -w

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# modify shape properties in an Excel xlsx file.
#
# reverse('©'), May 2012, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

# Create a new workbook called simple.xls and add a worksheet
my $workbook = Excel::Writer::XLSX->new( 'shape2.xlsx' );

die "Couldn't create new Excel file: $!.\n" unless defined $workbook;

my $worksheet = $workbook->add_worksheet();
$worksheet->hide_gridlines(2);

my $plain = $workbook->add_shape( 
    type => 'smileyFace', 
    text=> "Plain", 
    width=> 100, 
    height => 100,
);

my $bbformat = $workbook->add_format(color => 'red');
$bbformat->set_bold();
$bbformat->set_underline();
$bbformat->set_italic();

my $decor = $workbook->add_shape( 
    type => 'smileyFace', 
    text=> "Decorated", 
    rot => 45,
    width=> 200, 
    height => 100,
    format => $bbformat,
    typeface => 'Lucida Calligraphy',
    line_type => 'sysDot',
    line_weight => 3,
    fill => 'FFFF00',
    line => '3366FF',

);

$worksheet->insert_shape('A1', $plain,  50, 50);
$worksheet->insert_shape('A1', $decor, 250, 50);
$workbook->close();

__END__
C:\site\git\excel-writer-xlsx\examples\shape2.pl
