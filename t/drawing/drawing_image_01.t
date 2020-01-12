###############################################################################
#
# Tests for Excel::Writer::XLSX::Drawing methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Drawing;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $drawing = _new_object( \$got, 'Excel::Writer::XLSX::Drawing' );


###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tDrawing: _assemble_xml_file()";

my @dimensions = ( 2, 1, 0, 0, 3, 6, 533257, 190357, 1219200, 190500 );

my $drawing_object = $drawing->_add_drawing_object();

$drawing_object->{_type}        = 2;
$drawing_object->{_dimensions}  = \@dimensions;
$drawing_object->{_width}       = 1142857;
$drawing_object->{_height}      = 1142857;
$drawing_object->{_description} = 'republic.png';
$drawing_object->{_shape}       = undef;
$drawing_object->{_anchor}      = 2;
$drawing_object->{_rel_index}   = 1;

$drawing->{_embedded} = 1;

$drawing->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>2</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>3</xdr:col>
      <xdr:colOff>533257</xdr:colOff>
      <xdr:row>6</xdr:row>
      <xdr:rowOff>190357</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Picture 1" descr="republic.png"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="1219200" y="190500"/>
          <a:ext cx="1142857" cy="1142857"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
