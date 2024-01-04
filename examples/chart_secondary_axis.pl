#!/usr/bin/perl

#######################################################################
#
# A demo of a Line chart with a secondary axis in Excel::Writer::XLSX.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'chart_secondary_axis.xlsx' );
my $worksheet = $workbook->add_worksheet();
my $bold      = $workbook->add_format( bold => 1 );

# Add the worksheet data that the charts will refer to.
my $headings = [ 'Aliens', 'Humans', ];
my $data = [
    [ 2,  3,  4,  5,  6,  7 ],
    [ 10, 40, 50, 20, 10, 50 ],

];


$worksheet->write( 'A1', $headings, $bold );
$worksheet->write( 'A2', $data );

# Create a new chart object. In this case an embedded chart.
my $chart = $workbook->add_chart( type => 'line', embedded => 1 );

# Configure a series with a secondary axis
$chart->add_series(
    name    => '=Sheet1!$A$1',
    values  => '=Sheet1!$A$2:$A$7',
    y2_axis => 1,
);

$chart->add_series(
    name   => '=Sheet1!$B$1',
    values => '=Sheet1!$B$2:$B$7',
);

$chart->set_legend( position => 'right' );

# Add a chart title and some axis labels.
$chart->set_title( name => 'Survey results' );
$chart->set_x_axis( name => 'Days', );
$chart->set_y_axis( name => 'Population', major_gridlines => { visible => 0 } );
$chart->set_y2_axis( name => 'Laser wounds' );

# Insert the chart into the worksheet (with an offset).
$worksheet->insert_chart( 'D2', $chart, { x_offset => 25, y_offset => 10 } );

$workbook->close();

__END__
