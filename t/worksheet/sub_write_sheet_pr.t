###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 3;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;


###############################################################################
#
# 1. Test the _write_sheet_pr() method.
#
$caption  = " \tWorksheet: _write_sheet_pr()";
$expected = '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>';

$worksheet = _new_worksheet(\$got);
$worksheet->{_fit_page} = 1;
$worksheet->_write_sheet_pr();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_sheet_pr() method.
#
$caption  = " \tWorksheet: _write_sheet_pr()";
$expected = '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>';

$worksheet = _new_worksheet(\$got);

# Mock up the color palette.
$worksheet->{_palette}->[2] = [ 0xff, 0x00, 0x00, 0x00 ];

$worksheet->set_tab_color( 'red' );


$worksheet->_write_sheet_pr();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_sheet_pr() method.
#
$caption  = " \tWorksheet: _write_sheet_pr()";
$expected = '<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>';

$worksheet = _new_worksheet(\$got);

# Mock up the color palette.
$worksheet->{_palette}->[2] = [ 0xff, 0x00, 0x00, 0x00 ];

$worksheet->{_fit_page} = 1;
$worksheet->set_tab_color( 'red' );
$worksheet->_write_sheet_pr();

is( $got, $expected, $caption );




__END__


