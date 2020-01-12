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
# Test the _write_cell_xfs() method.
#
$caption = " \tStyles: _write_cell_xfs()";
$expected =
'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>';

my @formats = ( Excel::Writer::XLSX::Format->new( 0, {}, has_font => 1 ) );
my $num_fonts = 1;

$style = _new_style(\$got);
$style->_set_style_properties( \@formats, $num_fonts );
$style->_write_cell_xfs();

is( $got, $expected, $caption );

__END__


