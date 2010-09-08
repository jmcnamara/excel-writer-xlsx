###############################################################################
#
# Tests for Excel::XLSX::Writer::Worksheet methods.
#
# reverse('�'), September 2010, John McNamara, jmcnamara@cpan.org
#

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

my $workbook  = Excel::XLSX::Writer->new( $tmp_fh );
my $worksheet = $workbook->add_worksheet();
my $writer = new XML::Writer( OUTPUT => $got_fh );

$worksheet->{_writer} = $writer;

###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tWorksheet: _assemble_xml_file()";

$worksheet->write('B3', 123);
$worksheet->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

use Test::Differences;
eq_or_diff( $got, $expected, $caption,  { context => 1 } );
#is_deeply( $got, $expected, $caption );

###############################################################################
#
# Utility functions used by tests.
#
sub _expected_to_aref {

    my @data;

    while ( <DATA> ) {
        next unless /\S/;
        chomp;
        s{/>$}{ />};
        s{^\s+}{};
        push @data, $_;
    }

    return \@data;
}

sub _got_to_aref {

    my $xml_str = shift;

    $xml_str =~ s/\n//;

    # Split the XML into chunks at element boundaries.
    my @data = split /(?<=>)(?=<)/, $xml_str;

    return \@data;
}

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" mc:Ignorable="mv" mc:PreserveAttributes="mv:*">
  <sheetPr published="0" enableFormatConditionsCalculation="0"/>
  <dimension ref="B3"/>
  <sheetViews>
    <sheetView tabSelected="1" view="pageLayout" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="13"/>
  <sheetData>
    <row r="3">
      <c r="B3">
        <v>123</v>
      </c>
    </row>
  </sheetData>
  <sheetCalcPr fullCalcOnLoad="1"/>
  <phoneticPr fontId="1" type="noConversion"/>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
  <pageSetup paperSize="0" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/>
  <extLst>
    <ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main">
      <mx:PLV Mode="1" OnePage="0" WScale="0"/>
    </ext>
  </extLst>
</worksheet>
