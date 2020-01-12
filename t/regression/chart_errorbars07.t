###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_compare_xlsx_files _is_deep_diff);
use strict;
use warnings;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $filename     = 'chart_errorbars07.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];

my $ignore_elements = {

    'xl/charts/chart1.xml' => [ '<c:formatCode', ],
};


###############################################################################
#
# Test the creation of an Excel::Writer::XLSX file with error bars.
#
use Excel::Writer::XLSX;

my $workbook    = Excel::Writer::XLSX->new( $got_filename );
my $worksheet   = $workbook->add_worksheet();
my $chart       = $workbook->add_chart( type => 'stock', embedded => 1 );
my $date_format = $workbook->add_format( num_format => 14 );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids} = [ 45470848, 45472768 ];

my $data = [

    [ '2007-01-01T', '2007-01-02T', '2007-01-03T', '2007-01-04T', '2007-01-05T' ],
    [ 27.2,  25.03, 19.05, 20.34, 18.5 ],
    [ 23.49, 19.55, 15.12, 17.84, 16.34 ],
    [ 25.45, 23.05, 17.32, 20.45, 17.34 ],

];

for my $row ( 0 .. 4 ) {
    $worksheet->write_date_time( $row, 0, $data->[0]->[$row], $date_format );
    $worksheet->write( $row, 1, $data->[1]->[$row] );
    $worksheet->write( $row, 2, $data->[2]->[$row] );
    $worksheet->write( $row, 3, $data->[3]->[$row] );

}

$worksheet->set_column( 'A:D', 11 );


$chart->add_series(
    categories   => '=Sheet1!$A$1:$A$5',
    values       => '=Sheet1!$B$1:$B$5',
    y_error_bars => { type => 'standard_error' },
);

$chart->add_series(
    categories   => '=Sheet1!$A$1:$A$5',
    values       => '=Sheet1!$C$1:$C$5',
    y_error_bars => { type => 'standard_error' },
);

$chart->add_series(
    categories   => '=Sheet1!$A$1:$A$5',
    values       => '=Sheet1!$D$1:$D$5',
    y_error_bars => { type => 'standard_error' },
);


$worksheet->insert_chart( 'E9', $chart );

$workbook->close();


###############################################################################
#
# Compare the generated and existing Excel files.
#

my ( $got, $expected, $caption ) = _compare_xlsx_files(

    $got_filename,
    $exp_filename,
    $ignore_members,
    $ignore_elements,
);

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Cleanup.
#
unlink $got_filename;

__END__



