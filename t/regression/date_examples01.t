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
my $filename     = 'date_examples01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = ['xl/calcChain.xml', '\[Content_Types\].xml', 'xl/_rels/workbook.xml.rels'];
my $ignore_elements = {};


###############################################################################
#
# Example spreadsheet used in the tutorial.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

$worksheet->set_column( 'A:A', 30 );    # For extra visibility.

my $number = 41333.5;

$worksheet->write( 'A1', $number );             #   413333.5

my $format2 = $workbook->add_format( num_format => 'dd/mm/yy' );
$worksheet->write( 'A2', $number, $format2 );    #  28/02/13

my $format3 = $workbook->add_format( num_format => 'mm/dd/yy' );
$worksheet->write( 'A3', $number, $format3 );    #  02/28/13

my $format4 = $workbook->add_format( num_format => 'd\\-m\\-yyyy' );
$worksheet->write( 'A4', $number, $format4 );    #  28-2-2013

my $format5 = $workbook->add_format( num_format => 'dd/mm/yy\\ hh:mm' );
$worksheet->write( 'A5', $number, $format5 );    #  28/02/13 12:00

my $format6 = $workbook->add_format( num_format => 'd\\ mmm\\ yyyy' );
$worksheet->write( 'A6', $number, $format6 );    # 28 Feb 2013

my $format7 = $workbook->add_format( num_format => 'mmm\\ d\\ yyyy\\ hh:mm\\ AM/PM' );
$worksheet->write('A7', $number , $format7);     #  Feb 28 2013 12:00 PM




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



