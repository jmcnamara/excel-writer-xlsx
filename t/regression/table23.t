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
my $filename     = 'table23.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [ 'xl/calcChain.xml', '\[Content_Types\].xml', 'xl/_rels/workbook.xml.rels' ];
my $ignore_elements = {  'xl/workbook.xml' => ['<workbookView'] };


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with tables.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

# Set the column width to match the target worksheet.
$worksheet->set_column('B:F', 10.288);

# Write some strings to order the string table.
$worksheet->write_string('A1', 'Column1');
$worksheet->write_string('F1', 'Total');
$worksheet->write_string('B1', "Column'");
$worksheet->write_string('C1', 'Column#');
$worksheet->write_string('D1', 'Column[');
$worksheet->write_string('E1', 'Column]');

# Populate the data range.
# my $data =  [ 0, 0, 0, undef, undef, 0, 0, 0, 0, 0];
# $worksheet->write_row('B4', $data);
# $worksheet->write_row('B5', $data);


# Add the table.
$worksheet->add_table(
    'B3:F9',
    {
        total_row => 1,
        columns => [
            { header => 'Column1', total_string => 'Total' },
            { header => "Column'", total_function => 'sum' },
            { header => 'Column#', total_function => 'sum' },
            { header => 'Column[', total_function => 'sum' },
            { header => 'Column]', total_function => 'sum' },
          ],
    }
);



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
