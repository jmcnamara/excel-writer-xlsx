#!/usr/bin/perl

#######################################################################
#
# A demo of a Line chart in Excel::Writer::XLSX.
#
# reverse('©'), March 2011, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'chart_line.xlsx' );
my $worksheet = $workbook->add_worksheet();
my $bold      = $workbook->add_format( bold => 1 );

# Add the worksheet data that the charts will refer to.
my $headings = [ 'Number', 'Batch 1', 'Batch 2' ];
my $data = [
    [ 2, 3, 4, 5, 6, 7 ],
    [ 10, 40, 50, 20, 10, 50 ],
    [ 30, 60, 70, 50, 40, 30 ],

];

$worksheet->write( 'A1', $headings, $bold );
$worksheet->write( 'A2', $data );

# Create a new chart object. In this case an embedded chart.
my $chart = $workbook->add_chart( type => 'line', embedded => 1 );

# Configure the first series.
$chart->add_series(
    name       => '=Sheet1!$B$1',
    categories => '=Sheet1!$A$2:$A$7',
    values     => '=Sheet1!$B$2:$B$7',
);

# Configure the second series. Note the use of the array ref to define
# ranges: [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
$chart->add_series(
    name       => '=Sheet1!$C$1',
    categories => [ 'Sheet1', 1, 6, 0, 0 ],
    values     => [ 'Sheet1', 1, 6, 2, 2 ],
);

# Add a chart title and some axis labels.
$chart->set_title ( name => 'Results of sample analysis' );
$chart->set_x_axis( name => 'Test number' );
$chart->set_y_axis( name => 'Sample length (mm)' );

# Set an Excel chart style. Colors with white outline and shadow.
$chart->set_style( 10 );

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'D2', $chart, 25, 10 );

__END__
