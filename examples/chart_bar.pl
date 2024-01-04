#!/usr/bin/perl

#######################################################################
#
# A demo of an Bar chart in Excel::Writer::XLSX.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'chart_bar.xlsx' );
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

# Create a new chart object. In this case an embedded chart.
my $chart1 = $workbook->add_chart( type => 'bar', embedded => 1 );

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
$chart1->set_title ( name => 'Results of sample analysis' );
$chart1->set_x_axis( name => 'Test number' );
$chart1->set_y_axis( name => 'Sample length (mm)' );

# Set an Excel chart style. Blue colors with white outline and shadow.
$chart1->set_style( 11 );

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'D2', $chart1, { x_offset => 25, y_offset => 10 } );


#
# Create a stacked chart sub-type
#
my $chart2 = $workbook->add_chart(
    type     => 'bar',
    embedded => 1,
    subtype  => 'stacked'
);

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
$chart2->set_title ( name => 'Stacked Chart' );
$chart2->set_x_axis( name => 'Test number' );
$chart2->set_y_axis( name => 'Sample length (mm)' );

# Set an Excel chart style. Blue colors with white outline and shadow.
$chart2->set_style( 12 );

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'D18', $chart2, { x_offset => 25, y_offset => 10 } );


#
# Create a percent stacked chart sub-type
#
my $chart3 = $workbook->add_chart(
    type     => 'bar',
    embedded => 1,
    subtype  => 'percent_stacked'
);

# Configure the first series.
$chart3->add_series(
    name       => '=Sheet1!$B$1',
    categories => '=Sheet1!$A$2:$A$7',
    values     => '=Sheet1!$B$2:$B$7',
);

# Configure second series.
$chart3->add_series(
    name       => '=Sheet1!$C$1',
    categories => [ 'Sheet1', 1, 6, 0, 0 ],
    values     => [ 'Sheet1', 1, 6, 2, 2 ],
);

# Add a chart title and some axis labels.
$chart3->set_title ( name => 'Percent Stacked Chart' );
$chart3->set_x_axis( name => 'Test number' );
$chart3->set_y_axis( name => 'Sample length (mm)' );

# Set an Excel chart style. Blue colors with white outline and shadow.
$chart3->set_style( 13 );

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'D34', $chart3, { x_offset => 25, y_offset => 10 } );

$workbook->close();

__END__
