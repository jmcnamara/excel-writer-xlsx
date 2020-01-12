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
my $filename     = 'date_1904_02.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a Excel::Writer::XLSX file with date times in 1900 and
# 1904 epochs.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );

$workbook->set_1904();

my $worksheet = $workbook->add_worksheet();
my $format    = $workbook->add_format( num_format => 14 );

$worksheet->set_column( 'A:A', 12 );

$worksheet->write_date_time( 'A1', '1904-01-01T', $format );
$worksheet->write_date_time( 'A2', '1906-09-27T', $format );
$worksheet->write_date_time( 'A3', '1917-09-09T', $format );
$worksheet->write_date_time( 'A4', '1931-05-19T', $format );
$worksheet->write_date_time( 'A5', '2177-10-15T', $format );
$worksheet->write_date_time( 'A6', '4641-11-27T', $format );

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



