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
my $filename     = 'defined_name01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [
    qw(
      xl/printerSettings/printerSettings1.bin
      xl/worksheets/_rels/sheet1.xml.rels
      )
];

my $ignore_elements = {
    '[Content_Types].xml'      => ['<Default Extension="bin"'],
    'xl/worksheets/sheet1.xml' => ['<pageMargins', '<pageSetup'],
};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with defined names.
#
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( $got_filename );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet('Sheet 3');

$worksheet1->print_area( 'A1:E6' );
$worksheet1->autofilter('F1:G1');
$worksheet1->write( 'G1', 'Filter' );
$worksheet1->write( 'F1', 'Auto' );
$worksheet1->fit_to_pages( 2, 2 );

$workbook->define_name( q('Sheet 3'!Bar), q(='Sheet 3'!$A$1) );
$workbook->define_name( q(Abc),           q(=Sheet1!$A$1) );
$workbook->define_name( q(Baz),           q(=0.98) );
$workbook->define_name( q(Sheet1!Bar),    q(=Sheet1!$A$1) );
$workbook->define_name( q(Sheet2!Bar),    q(=Sheet2!$A$1) );
$workbook->define_name( q(Sheet2!aaa),    q(=Sheet2!$A$1) );
$workbook->define_name( q(_Egg),          q(=Sheet1!$A$1) );
$workbook->define_name( q(_Fog),          q(=Sheet1!$A$1) );

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



