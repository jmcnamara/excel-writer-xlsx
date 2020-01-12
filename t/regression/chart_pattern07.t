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
my $filename     = 'chart_pattern07.xlsx';
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
my $worksheet = $workbook->add_worksheet();
my $chart     = $workbook->add_chart( type => 'column', embedded => 1 );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids} = [ 110902272, 110756608 ];

my $data = [
    [ 2, 2, 2 ],
    [ 2, 2, 2 ],
    [ 2, 2, 2 ],
    [ 2, 2, 2 ],
    [ 2, 2, 2 ],
    [ 2, 2, 2 ],
    [ 2, 2, 2 ],
    [ 2, 2, 2 ],
];

$worksheet->write( 'A1', $data );

$chart->add_series(
    values  => '=Sheet1!$A$1:$A$3',
    pattern => {
        pattern  => 'percent_40',
        fg_color => '#C00000',
        bg_color => '#FFFFFF'
    }
);

$chart->add_series(
    values  => '=Sheet1!$B$1:$B$3',
    pattern => {
        pattern  => 'percent_90',
        fg_color => '#FF0000',
    }
);

$chart->add_series(
    values  => '=Sheet1!$C$1:$C$3',
    pattern => {
        pattern  => 'wide_upward_diagonal',
        fg_color => '#FFC000',
    }
);

$chart->add_series(
    values  => '=Sheet1!$D$1:$D$3',
    pattern => {
        pattern  => 'dark_horizontal',
        fg_color => '#FFFF00',
    }
);

$chart->add_series(
    values  => '=Sheet1!$E$1:$E$3',
    pattern => {
        pattern  => 'large_confetti',
        fg_color => '#92D050',
    }
);

$chart->add_series(
    values  => '=Sheet1!$F$1:$F$3',
    pattern => {
        pattern  => 'plaid',
        fg_color => '#00B050',
    }
);

$chart->add_series(
    values  => '=Sheet1!$G$1:$G$3',
    pattern => {
        pattern  => 'sphere',
        fg_color => '#00B0F0',
    }
);

$chart->add_series(
    values  => '=Sheet1!$H$1:$H$3',
    pattern => {
        pattern  => 'solid_diamond',
        fg_color => '#0070C0',
    }
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
