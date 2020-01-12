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
my $filename     = 'chart_clustered01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {

    # Ignore the page margins.
    'xl/charts/chart1.xml' => [ '<c:pageMargins' ],
};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $chart     = $workbook->add_chart(type => 'column', embedded => 1);


# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids} = [ 45886080, 45928832 ];

my $data = [
    [ 'Types',  'Sub Type',   'Value 1', 'Value 2', 'Value 3' ],
    [ 'Type 1', 'Sub Type A', 5000,      8000,      6000 ],
    [ '',       'Sub Type B', 2000,      3000,      4000 ],
    [ '',       'Sub Type C', 250,       1000,      2000 ],
    [ 'Type 2', 'Sub Type D', 6000,      6000,      6500 ],
    [ '',       'Sub Type E', 500,       300,       200 ],
];

my $cat_data = [
    [ 'Type 1',     undef,        undef,        'Type 2',     undef ],
    [ 'Sub Type A', 'Sub Type B', 'Sub Type C', 'Sub Type D', 'Sub Type E' ]
];


$worksheet->write_col( 'A1', $data );

$chart->add_series(
    name            => '=Sheet1!$C$1',
    categories      => '=Sheet1!$A$2:$B$6',
    values          => '=Sheet1!$C$2:$C$6',
    categories_data => $cat_data,
);

$chart->add_series(
    name       => '=Sheet1!$D$1',
    categories => '=Sheet1!$A$2:$B$6',
    values     => '=Sheet1!$D$2:$D$6',
);

$chart->add_series(
    name       => '=Sheet1!$E$1',
    categories => '=Sheet1!$A$2:$B$6',
    values     => '=Sheet1!$E$2:$E$6',
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
