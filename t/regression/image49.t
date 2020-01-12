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
my $filename     = 'image49.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with image(s).
#
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( $got_filename );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet();

$worksheet1->insert_image( 'A1',  $dir . 'images/blue.png' );
$worksheet1->insert_image( 'B3',  $dir . 'images/red.jpg' );
$worksheet1->insert_image( 'D5',  $dir . 'images/yellow.jpg' );
$worksheet1->insert_image( 'F9',  $dir . 'images/grey.png' );

$worksheet2->insert_image( 'A1',  $dir . 'images/blue.png' );
$worksheet2->insert_image( 'B3',  $dir . 'images/red.jpg' );
$worksheet2->insert_image( 'D5',  $dir . 'images/yellow.jpg' );
$worksheet2->insert_image( 'F9',  $dir . 'images/grey.png' );

$worksheet3->insert_image( 'A1',  $dir . 'images/blue.png' );
$worksheet3->insert_image( 'B3',  $dir . 'images/red.jpg' );
$worksheet3->insert_image( 'D5',  $dir . 'images/yellow.jpg' );
$worksheet3->insert_image( 'F9',  $dir . 'images/grey.png' );

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



