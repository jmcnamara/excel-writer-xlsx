###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::SharedStrings methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::SharedStrings;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Package::SharedStrings' );


###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tSharedStrings: _assemble_xml_file()";

$obj->_set_string_count(3);
$obj->_set_unique_count(3);
$obj->_add_strings(['abcdefg', '   abcdefg', 'abcdefg   ']);
$obj->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si>
    <t>abcdefg</t>
  </si>
  <si>
    <t xml:space="preserve">   abcdefg</t>
  </si>
  <si>
    <t xml:space="preserve">abcdefg   </t>
  </si>
</sst>
