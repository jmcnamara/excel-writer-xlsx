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
my $filename     = 'tutorial01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = ['xl/calcChain.xml', '\[Content_Types\].xml', 'xl/_rels/workbook.xml.rels'];
my $ignore_elements = {};


###############################################################################
#
# Example spreadsheet used in the tutorial.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

my @expenses = (
    ['Rent', 1000],
    ['Gas', 100],
    ['Food', 300],
    ['Gym', 50],
);

my $row = 0;

# Write the data to the worksheet.
for my $item (@expenses) {
    $worksheet->write($row, 0, $item->[0]);
    $worksheet->write($row, 1, $item->[1]);
    $row++
}

# Write the total.
$worksheet->write($row, 0, 'Total');
$worksheet->write($row, 1, '=SUM(B1:B4)', undef, 1450);


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



