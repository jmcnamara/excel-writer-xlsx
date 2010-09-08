###############################################################################
#
# Tests for Excel::XLSX::Writer::Workbook methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
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

use Test::Differences;
eq_or_diff( $got, $expected, $caption, { context => 1 } );

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
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
  <workbookPr date1904="1" showInkAnnotation="0" autoCompressPictures="0"/>
  <bookViews>
    <workbookView xWindow="-20" yWindow="-20" windowWidth="34400" windowHeight="20700" tabRatio="500"/>
  </bookViews>
  <sheets>
    <sheet name="Non Default Name" sheetId="1" r:id="rId1"/>
    <sheet name="Another Name" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="130000" concurrentCalc="0"/>
  <extLst>
    <ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main">
      <mx:ArchID Flags="2"/>
    </ext>
  </extLst>
</workbook>
