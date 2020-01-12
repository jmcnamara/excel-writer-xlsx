###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_style);
use strict;
use warnings;

use Test::More tests => 1;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $style;
my $tmp_fh;
my $tmp;
my $workbook;

open $tmp_fh, '>', \$tmp or die "Failed to open filehandle: $!";
$workbook = Excel::Writer::XLSX->new( $tmp_fh );


###############################################################################
#
# Test the _assemble_xml_file() method.
#
# Test for border colour styles.
#
$caption = " \tStyles: _assemble_xml_file()";

my $format1 = $workbook->add_format(
    left         => 1,
    right        => 1,
    top          => 1,
    bottom       => 1,
    diag_border  => 1,
    diag_type    => 3,
    left_color   => 'red',
    right_color  => 'red',
    top_color    => 'red',
    bottom_color => 'red',
    diag_color   => 'red'
);

$workbook->_set_default_xf_indices();
$workbook->_prepare_format_properties();

$style = _new_style( \$got );
$style->_set_style_properties(
    $workbook->{_xf_formats},
    $workbook->{_palette},
    $workbook->{_font_count},
    $workbook->{_num_format_count},
    $workbook->{_border_count},
    $workbook->{_fill_count},
    $workbook->{_custom_colors},
    $workbook->{_dxf_formats},
);
$style->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="2">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border diagonalUp="1" diagonalDown="1">
      <left style="thin">
        <color rgb="FFFF0000"/>
      </left>
      <right style="thin">
        <color rgb="FFFF0000"/>
      </right>
      <top style="thin">
        <color rgb="FFFF0000"/>
      </top>
      <bottom style="thin">
        <color rgb="FFFF0000"/>
      </bottom>
      <diagonal style="thin">
        <color rgb="FFFF0000"/>
      </diagonal>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
