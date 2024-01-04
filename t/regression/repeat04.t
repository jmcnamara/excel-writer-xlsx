###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
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
my $filename     = 'repeat04.xlsx';
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
    'xl/worksheets/sheet1.xml' => ['<pageMargins', '<pageSetup'],
};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with repeat rows
# when the sheet name contains a space.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet('Sheet 1');

$worksheet->repeat_rows( 0 );

$worksheet->write( 'A1', 'Foo' );

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



