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
my $filename     = 'hyperlink30.xlsx';
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

# Simulate custom colour for testing.
$workbook->{_custom_colors} = [ 'FF0000FF' ];

my $worksheet = $workbook->add_worksheet();
my $format1   = $workbook->add_format( hyperlink => 1 );
my $format2   = $workbook->add_format( color => 'red',  underline => 1 );
my $format3   = $workbook->add_format( color => 'blue', underline => 1 );

# Turn off default URL format for testing.
$worksheet->{_default_url_format} = undef;

$worksheet->write_url( 'A1', 'http://www.python.org/1', $format1 );
$worksheet->write_url( 'A2', 'http://www.python.org/2', $format2 );
$worksheet->write_url( 'A3', 'http://www.python.org/3', $format3 );

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



