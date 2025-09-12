###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# Copyright 2000-2025, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions '_new_style';
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


###############################################################################
#
# Test the _write_cell_styles() method.
#
$caption  = " \tStyles: _write_cell_styles()";
$expected = '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';

$style = _new_style(\$got);

$style->_write_cell_styles();

is( $got, $expected, $caption );

__END__


