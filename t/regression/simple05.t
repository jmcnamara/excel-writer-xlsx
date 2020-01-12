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
my $filename     = 'simple05.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test font formatting.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

$worksheet->set_row(5, 18);
$worksheet->set_row(6, 18);

my $format1 = $workbook->add_format( bold      => 1 );
my $format2 = $workbook->add_format( italic    => 1 );
my $format3 = $workbook->add_format( bold      => 1, italic => 1 );
my $format4 = $workbook->add_format( underline => 1 );
my $format5 = $workbook->add_format( font_strikeout => 1 );
my $format6 = $workbook->add_format( font_script => 1 );
my $format7 = $workbook->add_format( font_script => 2 );


$worksheet->write_string( 0, 0, 'Foo', $format1 );
$worksheet->write_string( 1, 0, 'Foo', $format2 );
$worksheet->write_string( 2, 0, 'Foo', $format3 );
$worksheet->write_string( 3, 0, 'Foo', $format4 );
$worksheet->write_string( 4, 0, 'Foo', $format5 );
$worksheet->write_string( 5, 0, 'Foo', $format6 );
$worksheet->write_string( 6, 0, 'Foo', $format7 );

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



