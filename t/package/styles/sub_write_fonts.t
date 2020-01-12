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
# Test the _write_fonts() method.
#
$caption = " \tStyles: _write_fonts()";
$expected =
'<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>';


my @formats = ( Excel::Writer::XLSX::Format->new( 0, {}, has_font => 1 ) );
my $num_fonts = 1;

$style = _new_style(\$got);

$style->_set_style_properties( \@formats, undef, $num_fonts );
$style->_write_fonts();

is( $got, $expected, $caption );

__END__


