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
my $filename     = 'optimize03.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );

$workbook->set_optimization();

my $worksheet = $workbook->add_worksheet();

my $bold = $workbook->add_format( bold => 1 );

$worksheet->set_column( 'A:A', 36, $bold );
$worksheet->set_column( 'B:B', 20 );
$worksheet->set_row( 0, 40 );

my $heading = $workbook->add_format(
    bold  => 1,
    color => 'blue',
    size  => 16,
    merge => 1,
    align => 'vcenter',
);

my $hyperlink_format = $workbook->add_format(
    color => 'blue',
    underline => 1,
);


my @headings = ( 'Features of Excel::Writer::XLSX', '' );
$worksheet->write_row( 'A1', \@headings, $heading );

my $text_format = $workbook->add_format(
    bold   => 1,
    italic => 1,
    color  => 'red',
    size   => 18,
    font   => 'Lucida Calligraphy'
);

$worksheet->write( 'A2', "Text" );
$worksheet->write( 'B2', "Hello Excel" );
$worksheet->write( 'A3', "Formatted text" );
$worksheet->write( 'B3', "Hello Excel", $text_format );
$worksheet->write( 'A4', "Unicode text" );
$worksheet->write( 'B4', "\x{0410} \x{0411} \x{0412} \x{0413} \x{0414}" );

my $num1_format = $workbook->add_format( num_format => '$#,##0.00' );
my $num2_format = $workbook->add_format( num_format => ' d mmmm yyy' );

$worksheet->write( 'A5', "Numbers" );
$worksheet->write( 'B5', 1234.56 );
$worksheet->write( 'A6', "Formatted numbers" );
$worksheet->write( 'B6', 1234.56, $num1_format );
$worksheet->write( 'A7', "Formatted numbers" );
$worksheet->write( 'B7', 37257, $num2_format );

$worksheet->set_selection( 'B8' );
$worksheet->write( 'A8', 'Formulas and functions, "=SIN(PI()/4)"' );
$worksheet->write( 'B8', '=SIN(PI()/4)' );

$worksheet->write( 'A9', "Hyperlinks" );
$worksheet->write( 'B9', 'http://www.perl.com/', $hyperlink_format );


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



