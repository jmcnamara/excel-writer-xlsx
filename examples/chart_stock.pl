#!/usr/bin/perl

#######################################################################
#
# A demo of a Stock chart in Excel::Writer::XLSX.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX;

my $workbook    = Excel::Writer::XLSX->new( 'chart_stock.xlsx' );
my $worksheet   = $workbook->add_worksheet();
my $bold        = $workbook->add_format( bold => 1 );
my $date_format = $workbook->add_format( num_format => 'dd/mm/yyyy' );
my $chart       = $workbook->add_chart( type => 'stock', embedded => 1 );


# Add the worksheet data that the charts will refer to.
my $headings = [ 'Date', 'High', 'Low', 'Close' ];
my $data = [

    [ '2007-01-01T', '2007-01-02T', '2007-01-03T', '2007-01-04T', '2007-01-05T' ],
    [ 27.2,  25.03, 19.05, 20.34, 18.5 ],
    [ 23.49, 19.55, 15.12, 17.84, 16.34 ],
    [ 25.45, 23.05, 17.32, 20.45, 17.34 ],

];

$worksheet->write( 'A1', $headings, $bold );

for my $row ( 0 .. 4 ) {
    $worksheet->write_date_time( $row+1, 0, $data->[0]->[$row], $date_format );
    $worksheet->write( $row+1, 1, $data->[1]->[$row] );
    $worksheet->write( $row+1, 2, $data->[2]->[$row] );
    $worksheet->write( $row+1, 3, $data->[3]->[$row] );

}

$worksheet->set_column( 'A:D', 11 );

# Add a series for each of the High-Low-Close columns.
$chart->add_series(
    categories => '=Sheet1!$A$2:$A$6',
    values     => '=Sheet1!$B$2:$B$6',
);

$chart->add_series(
    categories => '=Sheet1!$A$2:$A$6',
    values     => '=Sheet1!$C$2:$C$6',
);

$chart->add_series(
    categories => '=Sheet1!$A$2:$A$6',
    values     => '=Sheet1!$D$2:$D$6',
);

# Add a chart title and some axis labels.
$chart->set_title ( name => 'High-Low-Close', );
$chart->set_x_axis( name => 'Date', );
$chart->set_y_axis( name => 'Share price', );


$worksheet->insert_chart( 'E9', $chart );

$workbook->close();

__END__
