###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::VML methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions '_new_object';
use strict;
use warnings;
use Excel::Writer::XLSX::Package::VML;

use Test::More tests => 1;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $vml;


###############################################################################
#
# Test the _write_move_with_cells() method.
#
$caption  = " \tVML: _write_move_with_cells()";
$expected = '<x:MoveWithCells/>';

$vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );

$vml->_write_move_with_cells();

is( $got, $expected, $caption );

__END__


