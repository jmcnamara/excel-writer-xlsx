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
my $filename     = 'chart_bar13.xlsx';
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

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet1 = $workbook->add_worksheet();
my $chart1     = $workbook->add_chart( type => 'bar' );
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet();
my $chart2     = $workbook->add_chart( type => 'bar' );
my $worksheet4 = $workbook->add_worksheet();

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart1->{_chart}->{_axis_ids} = [ 40294272, 40295808 ];
$chart2->{_chart}->{_axis_ids} = [ 62356096, 62366080 ];

my $data = [
    [ 1, 2, 3, 4,  5 ],
    [ 2, 4, 6, 8,  10 ],
    [ 3, 6, 9, 12, 15 ],

];

$worksheet1->write( 'A1', $data );

$chart1->add_series( values => '=Sheet1!$A$1:$A$5' );
$chart1->add_series( values => '=Sheet1!$B$1:$B$5' );
$chart1->add_series( values => '=Sheet1!$C$1:$C$5' );

$chart2->add_series( values => '=Sheet1!$A$1:$A$5' );
$chart2->add_series( values => '=Sheet1!$B$1:$B$5' );
$chart2->add_series( values => '=Sheet1!$C$1:$C$5' );

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



