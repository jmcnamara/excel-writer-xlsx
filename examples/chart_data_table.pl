#!/usr/bin/perl

#######################################################################
#
# A demo of an Column chart with a data table on the X-axis using
# Excel::Writer::XLSX.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'chart_data_table.xlsx' );
my $worksheet = $workbook->add_worksheet();
my $bold      = $workbook->add_format( bold => 1 );

# Add the worksheet data that the charts will refer to.
my $headings = [ 'Number', 'Batch 1', 'Batch 2' ];
my $data = [
    [ 2,  3,  4,  5,  6,  7 ],
    [ 10, 40, 50, 20, 10, 50 ],
    [ 30, 60, 70, 50, 40, 30 ],

];

$worksheet->write( 'A1', $headings, $bold );
$worksheet->write( 'A2', $data );

# Create a column chart with a data table.
my $chart1 = $workbook->add_chart( type => 'column', embedded => 1 );

# Configure the first series.
$chart1->add_series(
    name       => '=Sheet1!$B$1',
    categories => '=Sheet1!$A$2:$A$7',
    values     => '=Sheet1!$B$2:$B$7',
);

# Configure second series. Note alternative use of array ref to define
# ranges: [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
$chart1->add_series(
    name       => '=Sheet1!$C$1',
    categories => [ 'Sheet1', 1, 6, 0, 0 ],
    values     => [ 'Sheet1', 1, 6, 2, 2 ],
);

# Add a chart title and some axis labels.
$chart1->set_title( name => 'Chart with Data Table' );
$chart1->set_x_axis( name => 'Test number' );
$chart1->set_y_axis( name => 'Sample length (mm)' );

# Set a default data table on the X-Axis.
$chart1->set_table();

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'D2', $chart1, { x_offset => 25, y_offset => 10 } );


#
# Create a second chart.
#
my $chart2 = $workbook->add_chart( type => 'column', embedded => 1 );

# Configure the first series.
$chart2->add_series(
    name       => '=Sheet1!$B$1',
    categories => '=Sheet1!$A$2:$A$7',
    values     => '=Sheet1!$B$2:$B$7',
);

# Configure second series.
$chart2->add_series(
    name       => '=Sheet1!$C$1',
    categories => [ 'Sheet1', 1, 6, 0, 0 ],
    values     => [ 'Sheet1', 1, 6, 2, 2 ],
);

# Add a chart title and some axis labels.
$chart2->set_title( name => 'Data Table with legend keys' );
$chart2->set_x_axis( name => 'Test number' );
$chart2->set_y_axis( name => 'Sample length (mm)' );

# Set a data table on the X-Axis with the legend keys showm.
$chart2->set_table( show_keys => 1 );

# Hide the chart legend since the keys are show on the data table.
$chart2->set_legend( position => 'none' );

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'D18', $chart2, { x_offset => 25, y_offset => 10 } );

$workbook->close();

__END__
