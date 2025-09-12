###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2025, John McNamara, jmcnamara@cpan.org
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
my $filename     = 'embed_image13.xlsx';
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

my $workbook  = Excel::Writer::XLSX->new( $got_filename );

my $worksheet1 = $workbook->add_worksheet();

$worksheet1->embed_image( 0, 0, $dir . 'images/red.png' );
$worksheet1->embed_image( 2, 0, $dir . 'images/blue.png' );
$worksheet1->embed_image( 4, 0, $dir . 'images/yellow.png' );

my $worksheet2 = $workbook->add_worksheet();

$worksheet2->embed_image( 0, 0, $dir . 'images/yellow.png' );
$worksheet2->embed_image( 2, 0, $dir . 'images/red.png' );
$worksheet2->embed_image( 4, 0, $dir . 'images/blue.png' );

my $worksheet3 = $workbook->add_worksheet();

$worksheet3->embed_image( 0, 0, $dir . 'images/blue.png' );
$worksheet3->embed_image( 2, 0, $dir . 'images/yellow.png' );
$worksheet3->embed_image( 4, 0, $dir . 'images/red.png' );


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



