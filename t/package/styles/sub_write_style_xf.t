###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
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
# Test the _write_style_xf() method.
#
$caption  = " \tStyles: _write_style_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>';

$style = _new_style(\$got);

$style->_write_style_xf(0, 0);

is( $got, $expected, $caption );

__END__


