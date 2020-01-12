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
my $filename     = 'optimize14.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with comments.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
$workbook->set_optimization();

my $worksheet = $workbook->add_worksheet();

$worksheet->write( 'A1',  'Foo' );
$worksheet->write( 'C7',  'Bar' );
$worksheet->write( 'G14', 'Baz' );

$worksheet->write_comment( 'A1',  'Some text' );
$worksheet->write_comment( 'D1',  'Some text' );
$worksheet->write_comment( 'C7',  'Some text' );
$worksheet->write_comment( 'E10', 'Some text' );
$worksheet->write_comment( 'G14', 'Some text' );

# Set the author to match the target XLSX file.
$worksheet->set_comments_author( 'John' );

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
