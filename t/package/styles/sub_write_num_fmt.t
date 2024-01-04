###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
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
# Test the _write_num_fmt() method.
#
$caption  = " \tStyles: _write_num_fmt()";
$expected = '<numFmt numFmtId="164" formatCode="#,##0.0"/>';

$style = _new_style(\$got);

$style->_write_num_fmt( 164, '#,##0.0' );

is( $got, $expected, $caption );

__END__


