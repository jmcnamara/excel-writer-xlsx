###############################################################################
#
# Tests for Excel::XLSX::Writer::Container::App methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::XLSX::Writer::Container::App;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $obj    = Excel::XLSX::Writer::Container::App->new();
my $writer = new XML::Writer( OUTPUT => $got_fh );

$obj->{_writer} = $writer;

###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tApp: _assemble_xml_file()";

$obj->_add_part_names('Sheet1');
$obj->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

is_deeply( $got, $expected, $caption );

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
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant>
        <vt:lpstr>Worksheets</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>1</vt:i4>
      </vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>Sheet1</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company>perl.org</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>12.0000</AppVersion>
</Properties>
