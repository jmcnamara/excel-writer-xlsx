###############################################################################
#
# Tests for Excel::Writer::XLSX::Drawing methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions '_new_object';
use strict;
use warnings;
use Excel::Writer::XLSX::Drawing;

use Test::More tests => 2;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $drawing;


###############################################################################
#
# Test the _write_c_nv_graphic_frame_pr() method.
#
$caption  = " \tDrawing: _write_c_nv_graphic_frame_pr()";
$expected = '<xdr:cNvGraphicFramePr/>';

$drawing = _new_object( \$got, 'Excel::Writer::XLSX::Drawing' );

$drawing->{_embedded} = 1;

$drawing->_write_c_nv_graphic_frame_pr();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_c_nv_graphic_frame_pr() method.
#
$caption  = " \tDrawing: _write_c_nv_graphic_frame_pr()";
$expected = '<xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>';

$drawing = _new_object( \$got, 'Excel::Writer::XLSX::Drawing' );

$drawing->_write_c_nv_graphic_frame_pr();

is( $got, $expected, $caption );


__END__


