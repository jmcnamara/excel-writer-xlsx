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
my $filename     = 'outline03.xlsx';
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
my $worksheet3 = $workbook->add_worksheet( 'Outline Columns' );

# Add a general format
my $bold = $workbook->add_format( bold => 1 );



# Example 3: Create a worksheet with outlined columns.
my $data = [
    [ 'Month', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Total' ],
    [ 'North', 50,    20,    15,    25,    65,    80,            ],
    [ 'South', 10,    20,    30,    50,    50,    50,            ],
    [ 'East',  45,    75,    50,    15,    75,    100,           ],
    [ 'West',  15,    15,    55,    35,    20,    50,            ],
];

# Add bold format to the first row
$worksheet3->set_row( 0, undef, $bold );

# Syntax: set_column($col1, $col2, $width, $XF, $hidden, $level, $collapsed)
$worksheet3->set_column( 'A:A', 10, $bold );
$worksheet3->set_column( 'B:G', 6, undef, 0, 1 );
$worksheet3->set_column( 'H:H', 10 );

# Write the data and a formula
$worksheet3->write_col( 'A1', $data );
$worksheet3->write( 'H2', '=SUM(B2:G2)', undef, 255 );
$worksheet3->write( 'H3', '=SUM(B3:G3)', undef, 210 );
$worksheet3->write( 'H4', '=SUM(B4:G4)', undef, 360 );
$worksheet3->write( 'H5', '=SUM(B5:G5)', undef, 190 );
$worksheet3->write( 'H6', '=SUM(H2:H5)', $bold, 1015 );

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



