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

use Test::More tests => 1;


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
# Test the _write_col_off() method.
#
$caption  = " \tDrawing: _write_col_off()";
$expected = '<xdr:colOff>457200</xdr:colOff>';

$drawing = _new_object( \$got, 'Excel::Writer::XLSX::Drawing' );

$drawing->_write_col_off( 457200 );

is( $got, $expected, $caption );

__END__


