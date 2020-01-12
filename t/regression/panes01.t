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
my $filename     = 'panes01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test an Excel::Writer::XLSX file with panes..
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet01 = $workbook->add_worksheet();
my $worksheet02 = $workbook->add_worksheet();
my $worksheet03 = $workbook->add_worksheet();
my $worksheet04 = $workbook->add_worksheet();
my $worksheet05 = $workbook->add_worksheet();
my $worksheet06 = $workbook->add_worksheet();
my $worksheet07 = $workbook->add_worksheet();
my $worksheet08 = $workbook->add_worksheet();
my $worksheet09 = $workbook->add_worksheet();
my $worksheet10 = $workbook->add_worksheet();
my $worksheet11 = $workbook->add_worksheet();
my $worksheet12 = $workbook->add_worksheet();
my $worksheet13 = $workbook->add_worksheet();

$worksheet01->write( 'A1', 'Foo' );
$worksheet02->write( 'A1', 'Foo' );
$worksheet03->write( 'A1', 'Foo' );
$worksheet04->write( 'A1', 'Foo' );
$worksheet05->write( 'A1', 'Foo' );
$worksheet06->write( 'A1', 'Foo' );
$worksheet07->write( 'A1', 'Foo' );
$worksheet08->write( 'A1', 'Foo' );
$worksheet09->write( 'A1', 'Foo' );
$worksheet10->write( 'A1', 'Foo' );
$worksheet11->write( 'A1', 'Foo' );
$worksheet12->write( 'A1', 'Foo' );
$worksheet13->write( 'A1', 'Foo' );

$worksheet01->freeze_panes( 'A2' );
$worksheet02->freeze_panes( 'A3' );
$worksheet03->freeze_panes( 'B1' );
$worksheet04->freeze_panes( 'C1' );
$worksheet05->freeze_panes( 'B2' );
$worksheet06->freeze_panes( 'G4' );
$worksheet07->freeze_panes( 3, 6, 3, 6, 1 );
$worksheet08->split_panes( 15, 0 );
$worksheet09->split_panes( 30, 0 );
$worksheet10->split_panes( 0,  8.46 );
$worksheet11->split_panes( 0,  17.57 );
$worksheet12->split_panes( 15, 8.46 );
$worksheet13->split_panes( 45, 54.14 );



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



