###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2024, John McNamara, jmcnamara@cpan.org
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
my $filename     = 'autofit12.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [ 'xl/calcChain.xml',
                       '\[Content_Types\].xml',
                       'xl/_rels/workbook.xml.rels' ];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

$worksheet->write_array_formula( 0, 0, 2, 0, '{=SUM(B1:C1*B2:C2)}',  undef, 1000  );


$worksheet->write(0, 1, 20);
$worksheet->write(1, 1, 30);
$worksheet->write(2, 1, 40);

$worksheet->write(0, 2, 10);
$worksheet->write(1, 2, 40);
$worksheet->write(2, 2, 20);

$worksheet->autofit();

# Put these after the autofit() so that the autofit in on the formula result.
$worksheet->write(1, 0, 1000);
$worksheet->write(2, 0, 1000);

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



