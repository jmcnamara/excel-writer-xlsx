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
$shape->set_line_weight(5);
$shape->set_line_type('lgDashDot');

my $drawing = _new_object( \$got, 'Excel::Writer::XLSX::Drawing' );
$drawing->{_embedded} = 1;

###############################################################################
#
# Test the _write_a_ln() method.
#
$caption = " \tDrawing: __write_a_ln() line weight and type";

$drawing->_add_drawing_object(
    3,     4,     8,     209550, 95250,  12, 22, 209660,
    96260, 10000, 20000, 95250,  190500, '', $shape
);

$drawing->_write_a_ln( $shape );

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<a:ln w="47625">
<a:solidFill>
<a:srgbClr val="000000"/>
</a:solidFill>
<a:prstDash val="lgDashDot"/>
<a:miter lim="800000"/>
<a:headEnd/>
<a:tailEnd/>
</a:ln>
