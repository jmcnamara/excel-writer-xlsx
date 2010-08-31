###############################################################################
#
# Tests for Excel::XLSX::Writer::Container::Core methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::XLSX::Writer::Container::Core;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $obj    = Excel::XLSX::Writer::Container::Core->new();
my $writer = new XML::Writer( OUTPUT => $got_fh );

$obj->{_writer} = $writer;

###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tCore: _assemble_xml_file()";

$obj->_set_creator('A User');
$obj->_set_modifier('Another User');
$obj->_set_creation_date('2010-01-02T00:00:00Z');
$obj->_set_modification_date('2010-01-03T00:00:00Z');
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
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>A User</dc:creator>
  <cp:lastModifiedBy>Another User</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2010-01-02T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2010-01-03T00:00:00Z</dcterms:modified>
</cp:coreProperties>
