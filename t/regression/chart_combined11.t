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
my $filename     = 'chart_combined11.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];

my $ignore_elements = { 'xl/charts/chart1.xml' => ['<c:dispBlanksAs'] };


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $chart_doughnut = $workbook->add_chart( type => 'doughnut', embedded => 1 );
my $chart_pie = $workbook->add_chart( type => 'pie', embedded => 1 );


$worksheet->write_col( 'H2', ['Donut', 25, 50, 25, 100] );
$worksheet->write_col( 'I2', ['Pie', 75, 1, 124] );

$chart_doughnut->add_series(
    name   => '=Sheet1!$H$2',
    values => '=Sheet1!$H$3:$H$6',
    points => [
        { fill => { color => '#FF0000' } },
        { fill => { color => '#FFC000' } },
        { fill => { color => '#00B050' } },
        { fill => { none  => 1 } },
    ],
);

$chart_doughnut->set_rotation( 270 );
$chart_doughnut->set_legend( none => 1 );
$chart_doughnut->set_chartarea(
    border => { none  => 1 },
    fill   => { none  => 1 },
);

$chart_pie->add_series(
    name   => '=Sheet1!$I$2',
    values => '=Sheet1!$I$3:$I$6',
    points => [
        { fill => { none  => 1 } },
        { fill => { color => '#FF0000' } },
        { fill => { none  => 1 } },
    ],
);

$chart_pie->set_rotation( 270 );

$chart_doughnut->combine($chart_pie);

$worksheet->insert_chart( 'A1', $chart_doughnut );

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



