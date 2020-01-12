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
use Excel::Writer::XLSX::Shape;
use Excel::Writer::XLSX::Drawing;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;

my $shape = Excel::Writer::XLSX::Shape->new();
# Set shape properties via AUTOLOAD
$shape->set_id(1000);
$shape->set_name('Connector 1');

# Set bulk shape properties via set_properties method
$shape->set_properties(type => 'straightConnector1', connect => 1);

my $drawing = _new_object( \$got, 'Excel::Writer::XLSX::Drawing' );
$drawing->{_embedded} = 1;

###############################################################################
#
# Test the _assemble_xml_file() method for connections
#
$caption = " \tDrawing: _assemble_xml_file() shape connection";

my @dimensions = ( 4, 8, 209550, 95250, 12, 22, 209660, 96260, 10000, 20000 );

my $drawing_object = $drawing->_add_drawing_object();

$drawing_object->{_type}        = 3;
$drawing_object->{_dimensions}  = \@dimensions;
$drawing_object->{_width}       = 95250;
$drawing_object->{_height}      = 190500;
$drawing_object->{_description} = undef;
$drawing_object->{_shape}       = $shape;
$drawing_object->{_anchor}      = 1;

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
            <xdr:colOff>209550</xdr:colOff>
            <xdr:row>8</xdr:row>
            <xdr:rowOff>95250</xdr:rowOff>
        </xdr:from>
        <xdr:to>
            <xdr:col>12</xdr:col>
            <xdr:colOff>209660</xdr:colOff>
            <xdr:row>22</xdr:row>
            <xdr:rowOff>96260</xdr:rowOff>
        </xdr:to>
        <xdr:cxnSp macro="">
            <xdr:nvCxnSpPr>
                <xdr:cNvPr id="1000" name="Connector 1"/>
                <xdr:cNvCxnSpPr>
                    <a:cxnSpLocks noChangeShapeType="1"/>
                </xdr:cNvCxnSpPr>
            </xdr:nvCxnSpPr>
            <xdr:spPr bwMode="auto">
                <a:xfrm>
                    <a:off x="10000" y="20000"/>
                    <a:ext cx="95250" cy="190500"/>
                </a:xfrm>
                <a:prstGeom prst="straightConnector1">
                    <a:avLst/>
                </a:prstGeom>
                <a:noFill/>
                <a:ln w="9525">
                    <a:solidFill>
                        <a:srgbClr val="000000"/>
                    </a:solidFill>
                    <a:round/>
                    <a:headEnd/>
                    <a:tailEnd/>
                </a:ln>
            </xdr:spPr>
        </xdr:cxnSp>
        <xdr:clientData/>
    </xdr:twoCellAnchor>
</xdr:wsDr>
