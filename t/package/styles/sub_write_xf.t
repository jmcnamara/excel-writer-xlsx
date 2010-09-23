###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# reverse('�'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_style';
use strict;
use warnings;

use Test::More tests => 5;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $style;
my %properties;
my $format;
my $index = 0;


###############################################################################
#
# 1. Test the _write_xf() method. Default properties.
#
%properties = ();

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />';

$format = Excel::Writer::XLSX::Format->new( $index, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_xf() method. Has font but is first XF.
#
%properties = ( has_font => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />';

$format = Excel::Writer::XLSX::Format->new( $index, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_xf() method. Has font but isn't first XF.
#
%properties = ( has_font => 1, font_index => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" />';

$format = Excel::Writer::XLSX::Format->new( $index, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_xf() method. Uses built-in number format.
#
%properties = ( num_format_index => 2 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />';

$format = Excel::Writer::XLSX::Format->new( $index, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_xf() method. Uses built-in number format + font.
#
%properties = (  num_format_index => 2, has_font => 1, font_index => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="2" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" />';

$format = Excel::Writer::XLSX::Format->new( $index, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );




__END__


