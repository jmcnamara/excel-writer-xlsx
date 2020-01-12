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

my @dimensions = ( 4, 8, 457200, 104775, 12, 22, 152400, 180975 );

my $drawing_object = $drawing->_add_drawing_object();

$drawing_object->{_type}        = 1;
$drawing_object->{_dimensions}  = \@dimensions;
$drawing_object->{_width}       = undef;
$drawing_object->{_height}      = undef;
$drawing_object->{_description} = undef;
$drawing_object->{_shape}       = undef;
$drawing_object->{_anchor}      = 1;
$drawing_object->{_rel_index}   = 1;


$drawing->{_embedded} = 1;

$drawing->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>4</xdr:col>
      <xdr:colOff>457200</xdr:colOff>
      <xdr:row>8</xdr:row>
      <xdr:rowOff>104775</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>12</xdr:col>
      <xdr:colOff>152400</xdr:colOff>
      <xdr:row>22</xdr:row>
      <xdr:rowOff>180975</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
