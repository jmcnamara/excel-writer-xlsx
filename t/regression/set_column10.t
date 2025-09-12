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
my $filename     = 'set_column01.xlsx';
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


$worksheet->set_column_pixels( "A:A",   1  );
$worksheet->set_column_pixels( "B:B",   2  );
$worksheet->set_column_pixels( "C:C",   3  );
$worksheet->set_column_pixels( "D:D",   4  );
$worksheet->set_column_pixels( "E:E",   5  );
$worksheet->set_column_pixels( "F:F",   6  );
$worksheet->set_column_pixels( "G:G",   7  );
$worksheet->set_column_pixels( "H:H",   8  );
$worksheet->set_column_pixels( "I:I",   9  );
$worksheet->set_column_pixels( "J:J",   10 );
$worksheet->set_column_pixels( "K:K",   11 );
$worksheet->set_column_pixels( "L:L",   12 );
$worksheet->set_column_pixels( "M:M",   13 );
$worksheet->set_column_pixels( "N:N",   14 );
$worksheet->set_column_pixels( "O:O",   15 );
$worksheet->set_column_pixels( "P:P",   16 );
$worksheet->set_column_pixels( "Q:Q",   17 );
$worksheet->set_column_pixels( "R:R",   18 );
$worksheet->set_column_pixels( "S:S",   19 );
$worksheet->set_column_pixels( "T:T",   20 );
$worksheet->set_column_pixels( "U:U",   21 );
$worksheet->set_column_pixels( "V:V",   22 );
$worksheet->set_column_pixels( "W:W",   23 );
$worksheet->set_column_pixels( "X:X",   24 );
$worksheet->set_column_pixels( "Y:Y",   25 );
$worksheet->set_column_pixels( "Z:Z",   26 );
$worksheet->set_column_pixels( "AB:AB", 65 );
$worksheet->set_column_pixels( "AC:AC", 66 );
$worksheet->set_column_pixels( "AD:AD", 67 );
$worksheet->set_column_pixels( "AE:AE", 68 );
$worksheet->set_column_pixels( "AF:AF", 69 );
$worksheet->set_column_pixels( "AG:AG", 70 );


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
