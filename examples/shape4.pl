#!/usr/bin/perl -w

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# demonstrate stenciling in an Excel xlsx file.
#
# reverse('©'), May 2012, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

# Create a new workbook called simple.xls and add a worksheet
my $workbook = Excel::Writer::XLSX->new( 'shape4.xlsx' );

die "Couldn't create new Excel file: $!.\n" unless defined $workbook;

my $worksheet = $workbook->add_worksheet();
$worksheet->hide_gridlines(2);

my $shape = $workbook->add_shape( 
    type => 'rect', 
    width=> 90, 
    height => 90,
);

for my $n (1..10) {
    # Change the last 5 rectangles to stars.  Previously inserted shapes stay as rectangles
    $shape->{type} = 'star5' if $n == 6;
    $shape->{text} = join (' ', $shape->{type}, $n); 
    $worksheet->insert_shape('A1', $shape,  $n * 100,  50);
}

$workbook->close();

__END__
C:\site\git\excel-writer-xlsx\examples\shape4.pl
