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
my $filename     = 'table12.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;


my $ignore_members  = [];
my $ignore_elements = {  'xl/workbook.xml' => ['<workbookView'] };


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with tables.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

my $data = [
    [ 'Foo', 1234, 2000 ],
    [ 'Bar', 1256, 4000 ],
    [ 'Baz', 2234, 3000 ],
];


# Set the column width to match the target worksheet.
$worksheet->set_column('C:F', 10.288);

# Add the table.
$worksheet->add_table( 'C2:F6', { data => $data } );

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



