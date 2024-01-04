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
# Test the _write_border() method.
#
$caption  = " \tStyles: _write_border()";
$expected = '<border><left/><right/><top/><bottom/><diagonal/></border>';

my $format = Excel::Writer::XLSX::Format->new( 0, {}, has_border => 1 );

$style = _new_style( \$got );
$style->_write_border( $format );

is( $got, $expected, $caption );

__END__


