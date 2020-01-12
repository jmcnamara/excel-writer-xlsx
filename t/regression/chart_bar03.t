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
my $filename     = 'chart_bar03.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];

my $ignore_elements = {

    # Ignore the page margins.
    'xl/charts/chart1.xml' => [ '<c:pageMargins' ],

    'xl/charts/chart2.xml' => [ '<c:pageMargins' ],

    # Ignore the workbookView.
    'xl/workbook.xml' => ['<workbookView'],

};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $chart1    = $workbook->add_chart( type => 'bar', embedded => 1 );
my $chart2    = $workbook->add_chart( type => 'bar', embedded => 1 );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart1->{_axis_ids} = [ 64265216, 64447616 ];
$chart2->{_axis_ids} = [ 86048128, 86058112 ];

my $data = [
    [ 1, 2, 3, 4,  5 ],
    [ 2, 4, 6, 8,  10 ],
    [ 3, 6, 9, 12, 15 ],

];

$worksheet->write( 'A1', $data );

$chart1->add_series(
    categories => '=Sheet1!$A$1:$A$5',
    values     => '=Sheet1!$B$1:$B$5',
);

$chart1->add_series(
    categories => '=Sheet1!$A$1:$A$5',
    values     => '=Sheet1!$C$1:$C$5',
);

$worksheet->insert_chart( 'E9', $chart1 );


$chart2->add_series(
    categories => '=Sheet1!$A$1:$A$4',
    values     => '=Sheet1!$B$1:$B$4',
);

$chart2->add_series(
    categories => '=Sheet1!$A$1:$A$4',
    values     => '=Sheet1!$C$1:$C$4',
);

$worksheet->insert_chart( 'F25', $chart2 );

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



