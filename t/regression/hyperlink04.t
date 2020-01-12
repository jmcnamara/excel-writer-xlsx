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
my $filename     = 'hyperlink04.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [];

my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with hyperlinks.
# This example doesn't have any link formatting and tests the relationship
# linkage code.
#
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( $got_filename );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet('Data Sheet');

# Turn off default URL format for testing.
$worksheet1->{_default_url_format} = undef;

$worksheet1->write_url( 'A1',   q(internal:Sheet2!A1) );
$worksheet1->write_url( 'A3',   q(internal:Sheet2!A1:A5) );
$worksheet1->write_url( 'A5',   q(internal:'Data Sheet'!D5), 'Some text' );
$worksheet1->write_url( 'E12',  q(internal:Sheet1!J1) );
$worksheet1->write_url( 'G17',  q(internal:Sheet2!A1), 'Some text', undef );
$worksheet1->write_url( 'A18',  q(internal:Sheet2!A1), undef, undef, 'Tool Tip 1' );
$worksheet1->write_url( 'A20',  q(internal:Sheet2!A1), 'More text', undef, 'Tool Tip 2' );

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



