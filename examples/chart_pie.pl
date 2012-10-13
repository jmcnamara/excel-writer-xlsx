#!/usr/bin/perl

#######################################################################
#
# A demo of a Pie chart in Excel::Writer::XLSX.
#
# reverse ('(c)'), March 2011, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'chart_pie.xlsx' );
my $worksheet = $workbook->add_worksheet();
my $bold      = $workbook->add_format( bold => 1 );

# Add the worksheet data that the charts will refer to.
my $headings = [ 'Category', 'Values' ];
my $data = [
    [ 'Apple', 'Cherry', 'Pecan' ],
    [ 60,       30,       10     ],
];

$worksheet->write( 'A1', $headings, $bold );
$worksheet->write( 'A2', $data );

# Create a new chart object. In this case an embedded chart.
my $chart = $workbook->add_chart( type => 'pie', embedded => 1 );

# Configure the series. Note the use of the array ref to define ranges:
# [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
$chart->add_series(
    name       => 'Pie sales data',
    categories => [ 'Sheet1', 1, 3, 0, 0 ],
    values     => [ 'Sheet1', 1, 3, 1, 1 ],
);

# Add a title.
$chart->set_title( name => 'Popular Pie Types' );

# Set an Excel chart style. Colors with white outline and shadow.
$chart->set_style( 10 );

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'C2', $chart, 25, 10 );

__END__
