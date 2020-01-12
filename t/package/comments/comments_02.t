###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Comments methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::Comments;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Package::Comments' );


###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tComments: _assemble_xml_file()";

$obj->_assemble_xml_file([ [ 1, 1, 'Some text', 'John', undef, 81, 'Calibri', 20, 2, [ 2, 0, 4, 4, 143, 10, 128, 74 ]] ] );

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>John</author>
  </authors>
  <commentList>
    <comment ref="B2" authorId="0">
      <text>
        <r>
          <rPr>
            <sz val="20"/>
            <color indexed="81"/>
            <rFont val="Calibri"/>
            <family val="2"/>
          </rPr>
          <t>Some text</t>
        </r>
      </text>
    </comment>
  </commentList>
</comments>
