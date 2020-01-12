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
my $filename     = 'custom_colors01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a Excel::Writer::XLSX file with custom colours.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

$workbook->set_custom_color( 40, '#26DA55' );
$workbook->set_custom_color( 41, '#792DC8' );
$workbook->set_custom_color( 42, '#646462' );

my $color1 = $workbook->add_format( bg_color => 40 );
my $color2 = $workbook->add_format( bg_color => 41 );
my $color3 = $workbook->add_format( bg_color => 42 );

$worksheet->write( 'A1', 'Foo', $color1 );
$worksheet->write( 'A2', 'Foo', $color2 );
$worksheet->write( 'A3', 'Foo', $color3 );

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



