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
$shape->set_flip_v(1);
$shape->set_flip_h(1);
$shape->set_rotation(90);

my $drawing = _new_object( \$got, 'Excel::Writer::XLSX::Drawing' );
$drawing->{_palette}  = $worksheet->{_palette};
$drawing->{_embedded} = 1;

###############################################################################
#
# Test the _write_a_xfrm() method.
#
$caption = " \tDrawing: _write_a_xfrm() shape rotation/flip";

$drawing->_add_drawing_object(
    3,     4,     8,     209550, 95250,  12, 22, 209660,
    96260, 10000, 20000, 95250,  190500, '', $shape
);

$drawing->_write_a_xfrm( 100, 200, 10, 20, $shape );

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<a:xfrm rot="5400000" flipH="1" flipV="1">
<a:off x="100" y="200"/>
<a:ext cx="10" cy="20"/>
</a:xfrm>
