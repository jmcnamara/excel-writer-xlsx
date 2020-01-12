###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Table methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::Table;
use Excel::Writer::XLSX::Worksheet;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Package::Table' );
my $worksheet = Excel::Writer::XLSX::Worksheet->new();

# Set the table properties.
$worksheet->add_table(
    'C5:D16',
    {
        banded_rows    => 0,
        first_column   => 1,
        last_column    => 1,
        banded_columns => 1
    }
);

$worksheet->_prepare_tables( 1, {} );


###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tTable: _assemble_xml_file()";

$obj->_set_properties( $worksheet->{_tables}->[0] );
$obj->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C5:D16" totalsRowShown="0">
  <autoFilter ref="C5:D16"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Column1"/>
    <tableColumn id="2" name="Column2"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="1" showLastColumn="1" showRowStripes="0" showColumnStripes="1"/>
</table>
