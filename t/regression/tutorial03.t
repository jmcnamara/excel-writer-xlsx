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
my $filename     = 'tutorial03.xlsx';
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

my $workbook     = Excel::Writer::XLSX->new( $got_filename );
my $worksheet    = $workbook->add_worksheet();
my $bold         = $workbook->add_format(bold => 1);
my $money_format = $workbook->add_format(num_format => '\\$#,##0');
my $date_format  = $workbook->add_format(num_format => 'mmmm\\ d\\ yyyy');

$worksheet->set_column('B:B', 15);

$worksheet->write('A1', 'Item', $bold);
$worksheet->write('B1', 'Date', $bold);
$worksheet->write('C1', 'Cost', $bold);

my @expenses = (
    [ 'Rent', '2013-01-13T', 1000 ],
    [ 'Gas',  '2013-01-14T', 100 ],
    [ 'Food', '2013-01-16T', 300 ],
    [ 'Gym',  '2013-01-20T', 50 ],
);

my $row = 1;

# Write the data to the worksheet.
for my $item (@expenses) {
    $worksheet->write_string($row,    0, $item->[0]);
    $worksheet->write_date_time($row, 1, $item->[1], $date_format);
    $worksheet->write_number($row,    2, $item->[2], $money_format);
    $row++
}

# Write the total.
$worksheet->write($row, 0, 'Total', $bold);
$worksheet->write($row, 2, '=SUM(C2:C5)', $money_format, 1450);


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



