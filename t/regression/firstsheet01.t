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
my $filename     = 'firstsheet01.xlsx';
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

my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet();
my $worksheet4 = $workbook->add_worksheet();
my $worksheet5 = $workbook->add_worksheet();
my $worksheet6 = $workbook->add_worksheet();
my $worksheet7 = $workbook->add_worksheet();
my $worksheet8 = $workbook->add_worksheet();
my $worksheet9 = $workbook->add_worksheet();
my $worksheet10 = $workbook->add_worksheet();
my $worksheet11 = $workbook->add_worksheet();
my $worksheet12 = $workbook->add_worksheet();
my $worksheet13 = $workbook->add_worksheet();
my $worksheet14 = $workbook->add_worksheet();
my $worksheet15 = $workbook->add_worksheet();
my $worksheet16 = $workbook->add_worksheet();
my $worksheet17 = $workbook->add_worksheet();
my $worksheet18 = $workbook->add_worksheet();
my $worksheet19 = $workbook->add_worksheet();
my $worksheet20 = $workbook->add_worksheet();


$worksheet8->set_first_sheet();
$worksheet20->activate();

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



