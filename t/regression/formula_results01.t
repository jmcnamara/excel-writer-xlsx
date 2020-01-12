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
my $filename     = 'formula_results01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [ 'xl/calcChain.xml',
                       '\[Content_Types\].xml',
                       'xl/_rels/workbook.xml.rels' ];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with formula errors.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

$worksheet->write_formula( 'A1',  '1+1',               undef, 2 );
$worksheet->write_formula( 'A2',  '"Foo"',             undef, 'Foo' );
$worksheet->write_formula( 'A3',  'IF(B3,FALSE,TRUE)', undef, 'TRUE' );
$worksheet->write_formula( 'A4',  'IF(B4,TRUE,FALSE)', undef, 'FALSE' );
$worksheet->write_formula( 'A5',  '#DIV/0!',           undef, '#DIV/0!' );
$worksheet->write_formula( 'A6',  '#N/A',              undef, '#N/A' );
$worksheet->write_formula( 'A7',  '#NAME?',            undef, '#NAME?' );
$worksheet->write_formula( 'A8',  '#NULL!',            undef, '#NULL!' );
$worksheet->write_formula( 'A9',  '#NUM!',             undef, '#NUM!' );
$worksheet->write_formula( 'A10', '#REF!',             undef, '#REF!' );
$worksheet->write_formula( 'A11', '#VALUE!',           undef, '#VALUE!' );
$worksheet->write_formula( 'A12', '1/0',               undef, '#DIV/0!' );


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
