###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2024, John McNamara, jmcnamara@cpan.org
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
my $filename     = 'chart_line08.xlsx';
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
my $chart    = $workbook->add_chart( type => 'line', embedded => 1 );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids} = [ 77034624, 77036544 ];
$chart->{_axis2_ids} = [ 95388032, 103040896 ];

my $data = [
    [ 1,  2,  3,  4,  5  ],
    [ 10, 40, 50, 20, 10 ],
    [ 1,  2,  3,  4,  5,  6,  7  ],
    [ 30, 10, 20, 40, 30, 10, 20 ],

];

$worksheet->write( 'A1', $data );


$chart->add_series(
    categories => '=Sheet1!$A$1:$A$5',
    values     => '=Sheet1!$B$1:$B$5',
);

$chart->add_series(
    categories => '=Sheet1!$C$1:$C$7',
    values     => '=Sheet1!$D$1:$D$7',
    y2_axis => 1,
);


$chart->set_x2_axis( label_position => 'next_to', visible => 1,  position => 'top' );
$chart->set_y2_axis( crossing => 'max' );

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
