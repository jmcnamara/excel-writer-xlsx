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
# Tests for diagonal border styles.
#
$caption = " \tStyles: _assemble_xml_file()";

my $format1 = $workbook->add_format( left      => 1 );
my $format2 = $workbook->add_format( right     => 1 );
my $format3 = $workbook->add_format( top       => 1 );
my $format4 = $workbook->add_format( bottom    => 1 );
my $format5 = $workbook->add_format( diag_type => 1, diag_border => 1 );
my $format6 = $workbook->add_format( diag_type => 2, diag_border => 1 );
my $format7 = $workbook->add_format( diag_type => 3 );  # Test default border.

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
  <borders count="8">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left style="thin">
        <color auto="1"/>
      </left>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right style="thin">
        <color auto="1"/>
      </right>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="thin">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style="thin">
        <color auto="1"/>
      </bottom>
      <diagonal/>
    </border>
    <border diagonalUp="1">
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal style="thin">
        <color auto="1"/>
      </diagonal>
    </border>
    <border diagonalDown="1">
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal style="thin">
        <color auto="1"/>
      </diagonal>
    </border>
    <border diagonalUp="1" diagonalDown="1">
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal style="thin">
        <color auto="1"/>
      </diagonal>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
