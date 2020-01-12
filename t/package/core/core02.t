###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Core methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::Core;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Package::Core' );

my %properties = (
    title    => 'This is an example spreadsheet',
    subject  => 'With document properties',
    author   => 'John McNamara',
    manager  => 'Dr. Heinz Doofenshmirtz',
    company  => 'of Wolves',
    category => 'Example spreadsheets',
    keywords => 'Sample, Example, Properties',
    comments => 'Created with Perl and Excel::Writer::XLSX',
    status   => 'Quo',
    created  => [ 15, 45, 19, 6, 3, 111 ],
);

###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tCore: _assemble_xml_file()";

$obj->_set_properties( \%properties );
$obj->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>This is an example spreadsheet</dc:title>
  <dc:subject>With document properties</dc:subject>
  <dc:creator>John McNamara</dc:creator>
  <cp:keywords>Sample, Example, Properties</cp:keywords>
  <dc:description>Created with Perl and Excel::Writer::XLSX</dc:description>
  <cp:lastModifiedBy>John McNamara</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2011-04-06T19:45:15Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2011-04-06T19:45:15Z</dcterms:modified>
  <cp:category>Example spreadsheets</cp:category>
  <cp:contentStatus>Quo</cp:contentStatus>
</cp:coreProperties>
