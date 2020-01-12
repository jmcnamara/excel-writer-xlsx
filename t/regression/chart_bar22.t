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
my $filename     = 'chart_bar22.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];

my $ignore_elements = {

    'xl/charts/chart1.xml' => [

        '<c:pageMargins',
    ],

};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $chart     = $workbook->add_chart( type => 'bar', embedded => 1 );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids} = [ 43706240, 43727104 ];


my $headers = [ 'Series 1', 'Series 2', 'Series 3' ];

my $data = [
    [ 'Category 1', 'Category 2', 'Category 3', 'Category 4' ],
    [ 4.3,          2.5,          3.5,          4.5 ],
    [ 2.4,          4.5,          1.8,          2.8 ],
    [ 2,            2,            3,            5 ],
];

$worksheet->set_column( 'A:D', 11 );

$worksheet->write( 'B1', $headers );
$worksheet->write( 'A2', $data );

$chart->add_series(
    categories      => '=Sheet1!$A$2:$A$5',
    values          => '=Sheet1!$B$2:$B$5',
    categories_data => $data->[0],
    values_data     => $data->[1],
);

$chart->add_series(
    categories      => '=Sheet1!$A$2:$A$5',
    values          => '=Sheet1!$C$2:$C$5',
    categories_data => $data->[0],
    values_data     => $data->[2],
);

$chart->add_series(
    categories      => '=Sheet1!$A$2:$A$5',
    values          => '=Sheet1!$D$2:$D$5',
    categories_data => $data->[0],
    values_data     => $data->[3],
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



