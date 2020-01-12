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
my $filename     = 'chart_bar14.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];

my $ignore_elements = {
    'xl/charts/chart1.xml' => ['<c:pageMargins'],
    'xl/charts/chart2.xml' => ['<c:pageMargins'],
    'xl/charts/chart3.xml' => ['<c:pageMargins'],
};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet();
my $chart1     = $workbook->add_chart( type => 'bar', embedded => 1 );
my $chart2     = $workbook->add_chart( type => 'bar', embedded => 1 );
my $chart3     = $workbook->add_chart( type => 'column' );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart1->{_axis_ids}           = [ 40294272, 40295808 ];
$chart2->{_axis_ids}           = [ 40261504, 65749760 ];
$chart3->{_chart}->{_axis_ids} = [ 65465728, 66388352 ];


my $data = [
    [ 1, 2, 3, 4,  5 ],
    [ 2, 4, 6, 8,  10 ],
    [ 3, 6, 9, 12, 15 ],

];

# Turn off default URL format for testing.
$worksheet2->{_default_url_format} = undef;

$worksheet2->write( 'A1', $data );
$worksheet2->write( 'A6', 'http://www.perl.com/' );

$chart3->add_series( values => '=Sheet2!$A$1:$A$5' );
$chart3->add_series( values => '=Sheet2!$B$1:$B$5' );
$chart3->add_series( values => '=Sheet2!$C$1:$C$5' );

$chart1->add_series( values => '=Sheet2!$A$1:$A$5' );
$chart1->add_series( values => '=Sheet2!$B$1:$B$5' );
$chart1->add_series( values => '=Sheet2!$C$1:$C$5' );

$chart2->add_series( values => '=Sheet2!$A$1:$A$5' );

$worksheet2->insert_chart( 'E9',  $chart1 );
$worksheet2->insert_chart( 'F25', $chart2 );

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



