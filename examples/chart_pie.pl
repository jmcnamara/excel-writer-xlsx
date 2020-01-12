#!/usr/bin/perl

#######################################################################
#
# A demo of a Pie chart in Excel::Writer::XLSX.
#
# The demo also shows how to set segment colours. It is possible to define
# chart colors for most types of Excel::Writer::XLSX charts via the
# add_series() method. However, Pie and Doughtnut charts are a special case
# since each segment is represented as a point so it is necessary to assign
# formatting to each point in the series.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
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
my $chart1 = $workbook->add_chart( type => 'pie', embedded => 1 );

# Configure the series. Note the use of the array ref to define ranges:
# [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
# See below for an alternative syntax.
$chart1->add_series(
    name       => 'Pie sales data',
    categories => [ 'Sheet1', 1, 3, 0, 0 ],
    values     => [ 'Sheet1', 1, 3, 1, 1 ],
);

# Add a title.
$chart1->set_title( name => 'Popular Pie Types' );

# Set an Excel chart style. Colors with white outline and shadow.
$chart1->set_style( 10 );

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'C2', $chart1, { x_offset => 25, y_offset => 10 } );


#
# Create a Pie chart with user defined segment colors.
#

# Create an example Pie chart like above.
my $chart2 = $workbook->add_chart( type => 'pie', embedded => 1 );

# Configure the series and add user defined segment colours.
$chart2->add_series(
    name       => 'Pie sales data',
    categories => '=Sheet1!$A$2:$A$4',
    values     => '=Sheet1!$B$2:$B$4',
    points     => [
        { fill => { color => '#5ABA10' } },
        { fill => { color => '#FE110E' } },
        { fill => { color => '#CA5C05' } },
    ],
);

# Add a title.
$chart2->set_title( name => 'Pie Chart with user defined colors' );


# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'C18', $chart2, { x_offset => 25, y_offset => 10 } );


#
# Create a Pie chart with rotation of the segments.
#

# Create an example Pie chart like above.
my $chart3 = $workbook->add_chart( type => 'pie', embedded => 1 );

# Configure the series.
$chart3->add_series(
    name       => 'Pie sales data',
    categories => '=Sheet1!$A$2:$A$4',
    values     => '=Sheet1!$B$2:$B$4',
);

# Add a title.
$chart3->set_title( name => 'Pie Chart with segment rotation' );

# Change the angle/rotation of the first segment.
$chart3->set_rotation(90);

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'C34', $chart3, { x_offset => 25, y_offset => 10 } );

$workbook->close();

__END__
