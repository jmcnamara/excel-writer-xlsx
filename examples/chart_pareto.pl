#!/usr/bin/perl

#######################################################################
#
# A demo of a Pareto chart in Excel::Writer::XLSX.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'chart_pareto.xlsx' );
my $worksheet = $workbook->add_worksheet();

# Formats used in the workbook.
my $bold           = $workbook->add_format( bold       => 1 );
my $percent_format = $workbook->add_format( num_format => '0.0%' );


# Widen the columns for visibility.
$worksheet->set_column( 'A:A', 15 );
$worksheet->set_column( 'B:C', 10 );

# Add the worksheet data that the charts will refer to.
my $headings = [ 'Reason', 'Number', 'Percentage' ];

my $reasons = [
    'Traffic',   'Child care', 'Public Transport', 'Weather',
    'Overslept', 'Emergency',
];

my $numbers  = [ 60,   40,    20,  15,  10,    5 ];
my $percents = [ 0.44, 0.667, 0.8, 0.9, 0.967, 1 ];

$worksheet->write_row( 'A1', $headings, $bold );
$worksheet->write_col( 'A2', $reasons );
$worksheet->write_col( 'B2', $numbers );
$worksheet->write_col( 'C2', $percents, $percent_format );


# Create a new column chart. This will be the primary chart.
my $column_chart = $workbook->add_chart( type => 'column', embedded => 1 );

# Add a series.
$column_chart->add_series(
    categories => '=Sheet1!$A$2:$A$7',
    values     => '=Sheet1!$B$2:$B$7',
);

# Add a chart title.
$column_chart->set_title( name => 'Reasons for lateness' );

# Turn off the chart legend.
$column_chart->set_legend( position => 'none' );

# Set the title and scale of the Y axes. Note, the secondary axis is set from
# the primary chart.
$column_chart->set_y_axis(
    name => 'Respondents (number)',
    min  => 0,
    max  => 120
);
$column_chart->set_y2_axis( max => 1 );

# Create a new line chart. This will be the secondary chart.
my $line_chart = $workbook->add_chart( type => 'line', embedded => 1 );

# Add a series, on the secondary axis.
$line_chart->add_series(
    categories => '=Sheet1!$A$2:$A$7',
    values     => '=Sheet1!$C$2:$C$7',
    marker     => { type => 'automatic' },
    y2_axis    => 1,
);


# Combine the charts.
$column_chart->combine( $line_chart );

# Insert the chart into the worksheet.
$worksheet->insert_chart( 'F2', $column_chart );

$workbook->close();

__END__
