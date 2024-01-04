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
my $filename     = 'set_column01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
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


$worksheet->set_column( "A:A",   0.08 );
$worksheet->set_column( "B:B",   0.17 );
$worksheet->set_column( "C:C",   0.25 );
$worksheet->set_column( "D:D",   0.33 );
$worksheet->set_column( "E:E",   0.42 );
$worksheet->set_column( "F:F",   0.5 );
$worksheet->set_column( "G:G",   0.58 );
$worksheet->set_column( "H:H",   0.67 );
$worksheet->set_column( "I:I",   0.75 );
$worksheet->set_column( "J:J",   0.83 );
$worksheet->set_column( "K:K",   0.92 );
$worksheet->set_column( "L:L",   1 );
$worksheet->set_column( "M:M",   1.14 );
$worksheet->set_column( "N:N",   1.29 );
$worksheet->set_column( "O:O",   1.43 );
$worksheet->set_column( "P:P",   1.57 );
$worksheet->set_column( "Q:Q",   1.71 );
$worksheet->set_column( "R:R",   1.86 );
$worksheet->set_column( "S:S",   2 );
$worksheet->set_column( "T:T",   2.14 );
$worksheet->set_column( "U:U",   2.29 );
$worksheet->set_column( "V:V",   2.43 );
$worksheet->set_column( "W:W",   2.57 );
$worksheet->set_column( "X:X",   2.71 );
$worksheet->set_column( "Y:Y",   2.86 );
$worksheet->set_column( "Z:Z",   3 );
$worksheet->set_column( "AB:AB", 8.57 );
$worksheet->set_column( "AC:AC", 8.71 );
$worksheet->set_column( "AD:AD", 8.86 );
$worksheet->set_column( "AE:AE", 9 );
$worksheet->set_column( "AF:AF", 9.14 );
$worksheet->set_column( "AG:AG", 9.29 );


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



