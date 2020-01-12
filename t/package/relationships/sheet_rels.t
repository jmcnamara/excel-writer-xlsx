###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Relationships methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::Relationships;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Package::Relationships' );


###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tRelationships: _assemble_xml_file()";

$obj->_add_worksheet_relationship( '/hyperlink', 'www.foo.com', 'External' );
$obj->_add_worksheet_relationship( '/hyperlink', 'link00.xlsx', 'External' );
$obj->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="www.foo.com" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="link00.xlsx" TargetMode="External"/>
</Relationships>
