###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2025, John McNamara, jmcnamara@cpan.org
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
my $filename     = 'table37.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with tables.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $chart     = $workbook->add_chart( type => 'column', embedded => 1 );

$worksheet->write(1, 0, 1);
$worksheet->write(2, 0, 2);
$worksheet->write(3, 0, 3);
$worksheet->write(4, 0, 4);
$worksheet->write(5, 0, 5);

$worksheet->write(1, 1, 10);
$worksheet->write(2, 1, 15);
$worksheet->write(3, 1, 20);
$worksheet->write(4, 1, 10);
$worksheet->write(5, 1, 15);

# Set the column width to match the target worksheet.
$worksheet->set_column('A:B', 10.288);

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids} = [ 88157568, 89138304 ];

$chart->add_series(
    name            => '=Sheet1!$B$1',
    categories      => '=Sheet1!$A$2:$A$6',
    values          => '=Sheet1!$B$2:$B$6',
);

$chart->set_title( none => 1 );

$worksheet->insert_chart( 'E9', $chart );

# Add the table.
$worksheet->add_table('A1:B6');


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



