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
my $filename     = 'set_row01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx2_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

$worksheet->set_row_pixels( 0,  1  );
$worksheet->set_row_pixels( 1,  2  );
$worksheet->set_row_pixels( 2,  3  );
$worksheet->set_row_pixels( 3,  4  );

$worksheet->set_row_pixels( 11, 12 );
$worksheet->set_row_pixels( 12, 13 );
$worksheet->set_row_pixels( 13, 14 );
$worksheet->set_row_pixels( 14, 15 );

$worksheet->set_row_pixels( 18, 19 );
$worksheet->set_row_pixels( 20, 21 );
$worksheet->set_row_pixels( 21, 22 );

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



