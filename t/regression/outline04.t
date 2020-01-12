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
my $filename     = 'outline04.xlsx';
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
my $worksheet4 = $workbook->add_worksheet( 'Outline levels' );

# Example 4: Show all possible outline levels.
#
my $levels = [
    "Level 1", "Level 2", "Level 3", "Level 4", "Level 5", "Level 6",
    "Level 7", "Level 6", "Level 5", "Level 4", "Level 3", "Level 2",
    "Level 1"
];

$worksheet4->write_col( 'A1', $levels );

$worksheet4->set_row( 0,  undef, undef, undef, 1 );
$worksheet4->set_row( 1,  undef, undef, undef, 2 );
$worksheet4->set_row( 2,  undef, undef, undef, 3 );
$worksheet4->set_row( 3,  undef, undef, undef, 4 );
$worksheet4->set_row( 4,  undef, undef, undef, 5 );
$worksheet4->set_row( 5,  undef, undef, undef, 6 );
$worksheet4->set_row( 6,  undef, undef, undef, 7 );
$worksheet4->set_row( 7,  undef, undef, undef, 6 );
$worksheet4->set_row( 8,  undef, undef, undef, 5 );
$worksheet4->set_row( 9,  undef, undef, undef, 4 );
$worksheet4->set_row( 10, undef, undef, undef, 3 );
$worksheet4->set_row( 11, undef, undef, undef, 2 );
$worksheet4->set_row( 12, undef, undef, undef, 1 );

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



