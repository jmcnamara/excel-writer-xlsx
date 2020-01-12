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
my $filename     = 'excel2003_style03.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [
    qw(
      xl/printerSettings/printerSettings1.bin
      xl/worksheets/_rels/sheet1.xml.rels
      )
];

my $ignore_elements = {
    '[Content_Types].xml'      => ['<Default Extension="bin"'],
    'xl/worksheets/sheet1.xml' => ['<pageMargins'],
};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename, {excel2003_style => 1});
my $worksheet = $workbook->add_worksheet();

$worksheet->set_paper(9);

$worksheet->set_header( 'Page &P' );
$worksheet->set_footer( '&A' );


my $bold = $workbook->add_format( bold => 1 );

$worksheet->write( 'A1', 'Foo' );
$worksheet->write( 'A2', 'Bar', $bold );


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
