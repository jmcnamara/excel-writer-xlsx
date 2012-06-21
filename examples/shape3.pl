#!/usr/bin/perl -w

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# scale shapes in an Excel xlsx file.
#
# reverse('©'), May 2012, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

# Create a new workbook called simple.xls and add a worksheet
my $workbook = Excel::Writer::XLSX->new( 'shape3.xlsx' );

die "Couldn't create new Excel file: $!.\n" unless defined $workbook;

my $worksheet = $workbook->add_worksheet();

my $normal = $workbook->add_shape( 
    name => 'chip', 
    type => 'diamond', 
    text=> "Normal", 
    width=> 100, 
    height => 100,
);

$worksheet->insert_shape('A1', $normal,  50, 50);
$normal->set_text('Scaled 3w x 2h');
$normal->set_name('Hope');
$worksheet->insert_shape('A1', $normal, 250, 50, 3, 2);
$workbook->close();

__END__
C:\site\git\excel-writer-xlsx\examples\shape3.pl
