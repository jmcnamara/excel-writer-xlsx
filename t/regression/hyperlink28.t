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

use Test::More tests => 3;

###############################################################################
#
# Tests setup.
#
my $filename     = 'hyperlink28.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [];

my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with hyperlinks.
# This example has link formatting.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $format    = $workbook->add_format( hyperlink => 1 );

$worksheet->write_url( 'A1', 'http://www.perl.org/', $format );

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
# Test2. Test with implicit hyperlink format.
#
$workbook  = Excel::Writer::XLSX->new( $got_filename );
$worksheet = $workbook->add_worksheet();

$worksheet->write_url( 'A1', 'http://www.perl.org/');

$workbook->close();


###############################################################################
#
# Compare the generated and existing Excel files.
#
( $got, $expected, $caption ) = _compare_xlsx_files(

    $got_filename,
    $exp_filename,
    $ignore_members,
    $ignore_elements,
);

_is_deep_diff( $got, $expected, $caption );




###############################################################################
#
# Test3. Test with the workbook default format.
#
$workbook  = Excel::Writer::XLSX->new( $got_filename );
$worksheet = $workbook->add_worksheet();
$format    = $workbook->get_default_url_format();

$worksheet->write_url( 'A1', 'http://www.perl.org/', $format);

$workbook->close();


###############################################################################
#
# Compare the generated and existing Excel files.
#
( $got, $expected, $caption ) = _compare_xlsx_files(

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



