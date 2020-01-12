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
# Test for border styles.
#
$caption = " \tStyles: _assemble_xml_file()";

my $format1  = $workbook->add_format( top => 7 );
my $format2  = $workbook->add_format( top => 4 );
my $format3  = $workbook->add_format( top => 11 );
my $format4  = $workbook->add_format( top => 9 );
my $format5  = $workbook->add_format( top => 3 );
my $format6  = $workbook->add_format( top => 1 );
my $format7  = $workbook->add_format( top => 12 );
my $format8  = $workbook->add_format( top => 13 );
my $format9  = $workbook->add_format( top => 10 );
my $format10 = $workbook->add_format( top => 8 );
my $format11 = $workbook->add_format( top => 2 );
my $format12 = $workbook->add_format( top => 5 );
my $format13 = $workbook->add_format( top => 6 );

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
  <borders count="14">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="hair">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dotted">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dashDotDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dashDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dashed">
        <color auto="1"/>
      </top>
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
      <top style="mediumDashDotDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="slantDashDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="mediumDashDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="mediumDashed">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="medium">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="thick">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="double">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="14">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="8" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="9" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="10" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="11" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="12" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="13" xfId="0" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
