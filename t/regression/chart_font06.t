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
my $filename     = 'chart_font06.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [];

my $ignore_elements = { 'xl/charts/chart1.xml' => ['<c:pageMargins'] };


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $chart     = $workbook->add_chart( type => 'bar', embedded => 1 );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids} = [ 49407488, 53740288 ];

my $data = [
    [ 1, 2, 3, 4,  5 ],
    [ 2, 4, 6, 8,  10 ],
    [ 3, 6, 9, 12, 15 ],

];

$worksheet->write( 'A1', $data );

$chart->add_series( values => '=Sheet1!$A$1:$A$5' );
$chart->add_series( values => '=Sheet1!$B$1:$B$5' );
$chart->add_series( values => '=Sheet1!$C$1:$C$5' );

$chart->set_title(
    name      => 'Title',
    name_font => {
        name         => 'Calibri',
        pitch_family => 34,
        charset      => 0,
        color        => 'yellow',
    },
);

$chart->set_x_axis(
    name      => 'XXX',
    name_font => {
        name         => 'Courier New',
        pitch_family => 49,
        charset      => 0,
        color        => '#92D050'
    },
    num_font => {
        name         => 'Arial',
        pitch_family => 34,
        charset      => 0,
        color        => '#00B0F0',
    },
);

$chart->set_y_axis(
    name      => 'YYY',
    name_font => {
        name         => 'Century',
        pitch_family => 18,
        charset      => 0,
        color        => 'red'
    },
    num_font => {
        bold      => 1,
        italic    => 1,
        underline => 1,
        color     => '#7030A0',
    },
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



