#!/usr/bin/perl

#######################################################################
#
# An example showing all 48 default chart styles available in Excel 2007
# using Excel::Writer::XLSX.. Note, these styles are not the same as the
# styles available in Excel 2013.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook = Excel::Writer::XLSX->new( 'chart_styles.xlsx' );

# Show the styles for all of these chart types.
my @chart_types = ( 'column', 'area', 'line', 'pie' );


for my $chart_type ( @chart_types ) {

    # Add a worksheet for each chart type.
    my $worksheet = $workbook->add_worksheet( ucfirst( $chart_type ) );
    $worksheet->set_zoom( 30 );
    my $style_number = 1;

    # Create 48 charts, each with a different style.
    for ( my $row_num = 0 ; $row_num < 90 ; $row_num += 15 ) {
        for ( my $col_num = 0 ; $col_num < 64 ; $col_num += 8 ) {

            my $chart = $workbook->add_chart(
                type     => $chart_type,
                embedded => 1
            );

            $chart->add_series( values => '=Data!$A$1:$A$6' );
            $chart->set_title( name => 'Style ' . $style_number );
            $chart->set_legend( none => 1 );
            $chart->set_style( $style_number );

            $worksheet->insert_chart( $row_num, $col_num, $chart );
            $style_number++;
        }
    }
}

# Create a worksheet with data for the charts.
my $data = [ 10, 40, 50, 20, 10, 50 ];
my $data_worksheet = $workbook->add_worksheet( 'Data' );
$data_worksheet->write_col( 'A1', $data );
$data_worksheet->hide();

$workbook->close();

__END__
