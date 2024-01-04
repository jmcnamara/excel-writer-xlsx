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
my $filename     = 'set_column06.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx2_$filename";
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
my $chart     = $workbook->add_chart( type => 'line', embedded => 1 );

my $bold   = $workbook->add_format( bold   => 1 );
my $italic = $workbook->add_format( italic => 1 );


# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids} = [ 69197824, 69199360 ];

my $data = [
    [ 1, 2, 3, 4,  5 ],
    [ 2, 4, 6, 8,  10 ],
    [ 3, 6, 9, 12, 15 ],

];

$worksheet->write( 'A1', 'Foo', $bold );
$worksheet->write( 'B1', 'Bar', $italic );
$worksheet->write( 'A2', $data );

$worksheet->set_row_pixels( 12, undef, undef, 1 );
$worksheet->set_column_pixels( 'F:F', undef, undef, 1 );


$chart->add_series( values => '=Sheet1!$A$2:$A$6' );
$chart->add_series( values => '=Sheet1!$B$2:$B$6' );
$chart->add_series( values => '=Sheet1!$C$2:$C$6' );

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



