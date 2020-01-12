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
my $filename     = 'chart_bar04.xlsx';
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

my $workbook   = Excel::Writer::XLSX->new( $got_filename );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $chart1     = $workbook->add_chart( type => 'bar', embedded => 1 );
my $chart2     = $workbook->add_chart( type => 'bar', embedded => 1 );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart1->{_axis_ids} = [ 64446848, 64448384 ];
$chart2->{_axis_ids} = [ 85389696, 85391232 ];

my $data = [
    [ 1, 2, 3, 4,  5 ],
    [ 2, 4, 6, 8,  10 ],
    [ 3, 6, 9, 12, 15 ],

];

# Sheet 1
$worksheet1->write( 'A1', $data );

$chart1->add_series(
    categories => '=Sheet1!$A$1:$A$5',
    values     => '=Sheet1!$B$1:$B$5',
);

$chart1->add_series(
    categories => '=Sheet1!$A$1:$A$5',
    values     => '=Sheet1!$C$1:$C$5',
);

$worksheet1->insert_chart( 'E9', $chart1 );

# Sheet 2
$worksheet2->write( 'A1', $data );

$chart2->add_series(
    categories => '=Sheet2!$A$1:$A$5',
    values     => '=Sheet2!$B$1:$B$5',
);

$chart2->add_series(
    categories => '=Sheet2!$A$1:$A$5',
    values     => '=Sheet2!$C$1:$C$5',
);

$worksheet2->insert_chart( 'E9', $chart2 );

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
