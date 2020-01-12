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
my $filename     = 'outline01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members =
  [ 'xl/calcChain.xml', '\[Content_Types\].xml', 'xl/_rels/workbook.xml.rels' ];
my $ignore_elements = { 'xl/workbook.xml' => ['<workbookView'] };


###############################################################################
#
# Test the creation of a outlines in a Excel::Writer::XLSX file. These tests
# are based on the outline programs in the examles directory.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet1 = $workbook->add_worksheet( 'Outlined Rows' );

# Add a general format
my $bold = $workbook->add_format( bold => 1 );

# For outlines the important parameters are $hidden and $level. Rows with the
# same $level are grouped together. The group will be collapsed if $hidden is
# non-zero. $height and $XF are assigned default values if they are undef.
#
# The syntax is: set_row($row, $height, $XF, $hidden, $level, $collapsed)
#
$worksheet1->set_row( 1, undef, undef, 0, 2 );
$worksheet1->set_row( 2, undef, undef, 0, 2 );
$worksheet1->set_row( 3, undef, undef, 0, 2 );
$worksheet1->set_row( 4, undef, undef, 0, 2 );
$worksheet1->set_row( 5, undef, undef, 0, 1 );

$worksheet1->set_row( 6,  undef, undef, 0, 2 );
$worksheet1->set_row( 7,  undef, undef, 0, 2 );
$worksheet1->set_row( 8,  undef, undef, 0, 2 );
$worksheet1->set_row( 9,  undef, undef, 0, 2 );
$worksheet1->set_row( 10, undef, undef, 0, 1 );


# Add a column format for clarity
$worksheet1->set_column( 'A:A', 20 );

# Add the data, labels and formulas
$worksheet1->write( 'A1', 'Region', $bold );
$worksheet1->write( 'A2', 'North' );
$worksheet1->write( 'A3', 'North' );
$worksheet1->write( 'A4', 'North' );
$worksheet1->write( 'A5', 'North' );
$worksheet1->write( 'A6', 'North Total', $bold );

$worksheet1->write( 'B1', 'Sales', $bold );
$worksheet1->write( 'B2', 1000 );
$worksheet1->write( 'B3', 1200 );
$worksheet1->write( 'B4', 900 );
$worksheet1->write( 'B5', 1200 );
$worksheet1->write( 'B6', '=SUBTOTAL(9,B2:B5)', $bold, 4300 );

$worksheet1->write( 'A7',  'South' );
$worksheet1->write( 'A8',  'South' );
$worksheet1->write( 'A9',  'South' );
$worksheet1->write( 'A10', 'South' );
$worksheet1->write( 'A11', 'South Total', $bold );

$worksheet1->write( 'B7',  400 );
$worksheet1->write( 'B8',  600 );
$worksheet1->write( 'B9',  500 );
$worksheet1->write( 'B10', 600 );
$worksheet1->write( 'B11', '=SUBTOTAL(9,B7:B10)', $bold, 2100 );

$worksheet1->write( 'A12', 'Grand Total',         $bold );
$worksheet1->write( 'B12', '=SUBTOTAL(9,B2:B10)', $bold, 6400 );


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



