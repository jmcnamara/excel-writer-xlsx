###############################################################################
#
# Tests for Excel::XLSX::Writer::Package::Relationships methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::XLSX::Writer::Package::Relationships;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $obj    = Excel::XLSX::Writer::Package::Relationships->new();
my $writer = new XML::Writer( OUTPUT => $got_fh );

$obj->{_writer} = $writer;

###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tRelationships: _assemble_xml_file()";

$obj->_add_document_relationship( '/worksheet',     'worksheets/sheet1' );
$obj->_add_document_relationship( '/theme',         'theme/theme1' );
$obj->_add_document_relationship( '/styles',        'styles' );
$obj->_add_document_relationship( '/sharedStrings', 'sharedStrings' );
$obj->_add_document_relationship( '/calcChain',     'calcChain' );
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
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>
