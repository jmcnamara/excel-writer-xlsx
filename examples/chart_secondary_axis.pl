#!/usr/bin/env perl

use strict;
use warnings;
use Excel::Writer::XLSX;

##my $path = '/mac_share/kablamo/secondary_axis.cpan.new.xlsx';
my $path      = 'secondary.new/secondary_axis.cpan.new.xlsx';
my $workbook  = Excel::Writer::XLSX->new( $path );
my $worksheet = $workbook->add_worksheet();
my $bold      = $workbook->add_format( bold => 1 );

# Add the worksheet data that the charts will refer to.
my $headings = [ 'Hungry aliens', 'Plump humans', ];
my $data = [
    [ 2,  3,  4,  5,  6,  7 ],
    [ 10, 40, 50, 20, 10, 50 ],
##  [ 30, 60, 70, 50, 40, 30 ],
];

$worksheet->write( 'A1', $headings, $bold );
$worksheet->write( 'A2', $data );

# Create a new chart object. In this case an embedded chart.
my $chart = $workbook->add_chart( type => 'line', embedded => 1 );

$chart->add_series(
    name    => '=Sheet1!$A$1',
    values  => '=Sheet1!$A$2:$A$7',
    y2_axis => 1,
);

# ranges: [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
$chart->add_series(
    name   => '=Sheet1!$B$1',
    values => '=Sheet1!$B$2:$B$7',
);


$chart->set_legend( position => 'right' );    # none, bottom, etc

# Add a chart title and some axis labels.
$chart->set_title( name => 'Intergalactic survey results' );
$chart->set_x_axis( name => 'Days', );
$chart->set_y_axis( name => 'Population', major_gridlines => 0 );
$chart->set_x2_axis();
$chart->set_y2_axis( name => 'not sure what this measures', );

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'D2', $chart, 25, 10 );
