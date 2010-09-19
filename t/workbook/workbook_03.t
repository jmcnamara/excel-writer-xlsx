###############################################################################
#
# Tests for Excel::XLSX::Writer::Workbook methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff);
use strict;
use warnings;
use Excel::XLSX::Writer;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $tmp_fh, '>', \my $tmp or die "Failed to open filehandle: $!";
open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $workbook   = Excel::XLSX::Writer->new( $tmp_fh );
my $worksheet1 = $workbook->add_worksheet('Non Default Name');
my $worksheet2 = $workbook->add_worksheet('Another Name');
my $writer     = new XML::Writer( OUTPUT => $got_fh );

$workbook->{_writer} = $writer;

###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tWorkbook: _assemble_xml_file()";

$workbook->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
  <workbookPr defaultThemeVersion="124226"/>
  <bookViews>
    <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
  </bookViews>
  <sheets>
    <sheet name="Non Default Name" sheetId="1" r:id="rId1"/>
    <sheet name="Another Name" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="124519"/>
</workbook>
