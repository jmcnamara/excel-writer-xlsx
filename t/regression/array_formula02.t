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
my $filename     = 'array_formula02.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [ 'xl/calcChain.xml', '\[Content_Types\].xml', 'xl/_rels/workbook.xml.rels' ];
my $ignore_elements = {  'xl/workbook.xml' => ['<workbookView'] };


###############################################################################
#
# Test the creation of an Excel::Writer::XLSX file with an array formula.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $format    = $workbook->add_format( bold => 1);

my $data = [ 0, 0, 0 ];

$worksheet->write_col( 'B1', $data );
$worksheet->write_col( 'C1', $data );

$worksheet->write_array_formula( 'A1:A3', '{=SUM(B1:C1*B2:C2)}',  $format, 0  );


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



