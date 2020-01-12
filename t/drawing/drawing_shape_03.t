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
my $worksheet = Excel::Writer::XLSX::Worksheet->new();

my $shape = Excel::Writer::XLSX::Shape->new();
$shape->set_id(1000);
$shape->set_start(1001);
$shape->set_start_index(1);
$shape->set_end(1002);
$shape->set_end_index(4);

my $drawing = _new_object( \$got, 'Excel::Writer::XLSX::Drawing' );
$drawing->{_palette}  = $worksheet->{_palette};
$drawing->{_embedded} = 1;

###############################################################################
#
# Test the _assemble_xml_file() method for shape connections.
#
$caption = " \tDrawing: _write_nv_cxn_sp_pr() shape connection";

$drawing->_add_drawing_object(
    3,     4,     8,     209550, 95250,  12, 22, 209660,
    96260, 10000, 20000, 95250,  190500, '', $shape, 1
);

$drawing->_write_nv_cxn_sp_pr( 1, $shape );

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<xdr:nvCxnSpPr>
<xdr:cNvPr id="1000" name="rect 1"/>
<xdr:cNvCxnSpPr>
<a:cxnSpLocks noChangeShapeType="1"/>
<a:stCxn id="1001" idx="1"/>
<a:endCxn id="1002" idx="4"/>
</xdr:cNvCxnSpPr>
</xdr:nvCxnSpPr>
