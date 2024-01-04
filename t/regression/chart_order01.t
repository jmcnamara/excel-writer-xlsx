###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
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
my $filename     = 'chart_order01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];

my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( $got_filename );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet();

my $chart1 = $workbook->add_chart(
    type     => 'column',
    embedded => 1,
);

my $chart2 = $workbook->add_chart(
    type     => 'bar',
    embedded => 1,
);

my $chart3 = $workbook->add_chart(
    type     => 'line',
    embedded => 1,
);

my $chart4 = $workbook->add_chart(
    type     => 'pie',
    embedded => 1,
);




# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart1->{_axis_ids} = [ 54976896, 54978432 ];
$chart2->{_axis_ids} = [ 54310784, 54312320 ];
$chart3->{_axis_ids} = [ 69816704, 69818240 ];
$chart4->{_axis_ids} = [ 69816704, 69818240 ];

my $data = [
    [ 1, 2, 3, 4,  5 ],
    [ 2, 4, 6, 8,  10 ],
    [ 3, 6, 9, 12, 15 ],

];

$worksheet1->write( 'A1', $data );
$worksheet2->write( 'A1', $data );
$worksheet3->write( 'A1', $data );

$chart1->add_series( values => '=Sheet1!$A$1:$A$5' );
$chart2->add_series( values => '=Sheet2!$A$1:$A$5' );
$chart3->add_series( values => '=Sheet3!$A$1:$A$5' );
$chart4->add_series( values => '=Sheet1!$B$1:$B$5' );


$worksheet1->insert_chart( 'E9',  $chart1 );
$worksheet2->insert_chart( 'E9',  $chart2 );
$worksheet3->insert_chart( 'E9',  $chart3 );
$worksheet1->insert_chart( 'E24', $chart4 );

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



