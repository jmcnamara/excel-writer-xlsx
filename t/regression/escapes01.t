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
my $filename     = 'escapes01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [ 'xl/calcChain.xml', '\[Content_Types\].xml', 'xl/_rels/workbook.xml.rels' ];
my $ignore_elements = { 'xl/workbook.xml' => ['<workbookView'] };


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with strings that
# require XML escaping.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet('5&4');

$worksheet->write_formula( 'A1', q{=IF(1>2,0,1)},            undef, 1 );
$worksheet->write_formula( 'A2', q{=CONCATENATE("'","<>&")}, undef, q{'<>&} );
$worksheet->write_formula( 'A3', q{=1&"b"},                  undef, q{1b} );
$worksheet->write_formula( 'A4', q{="'"},                    undef, q{'} );
$worksheet->write_formula( 'A5', q{=""""},                   undef, q{"} );
$worksheet->write_formula( 'A6', q{="&" & "&"},              undef, q{&&} );

$worksheet->write_string( 'A8', q{"&<>} );


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



