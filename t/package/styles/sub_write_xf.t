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

use Test::More tests => 36;


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


###############################################################################
#
# 1. Test the _write_xf() method. Default properties.
#
%properties = ();

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_xf() method. Has font but is first XF.
#
%properties = ( has_font => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_xf() method. Has font but isn't first XF.
#
%properties = ( has_font => 1, font_index => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_xf() method. Uses built-in number format.
#
%properties = ( num_format_index => 2 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_xf() method. Uses built-in number format + font.
#
%properties = (  num_format_index => 2, has_font => 1, font_index => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="2" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_xf() method. Vertical alignment = top.
#
%properties = ( align => 'top' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 7. Test the _write_xf() method. Vertical alignment = centre.
#
%properties = ( align => 'vcenter' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 8. Test the _write_xf() method. Vertical alignment = bottom.
#
%properties = ( align => 'bottom' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"/>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 9. Test the _write_xf() method. Vertical alignment = justify.
#
%properties = ( align => 'vjustify' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="justify"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 10. Test the _write_xf() method. Vertical alignment = distributed.
#
%properties = ( align => 'vdistributed' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="distributed"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 11. Test the _write_xf() method. Horizontal alignment = left.
#
%properties = ( align => 'left' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 12. Test the _write_xf() method. Horizontal alignment = center.
#
%properties = ( align => 'center' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 13. Test the _write_xf() method. Horizontal alignment = right.
#
%properties = ( align => 'right' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 14. Test the _write_xf() method. Horizontal alignment = left + indent.
#
%properties = ( align => 'left', indent => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 15. Test the _write_xf() method. Horizontal alignment = right + indent.
#
%properties = ( align => 'right', indent => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" indent="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 16. Test the _write_xf() method. Horizontal alignment = fill.
#
%properties = ( align => 'fill' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="fill"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 17. Test the _write_xf() method. Horizontal alignment = justify.
#
%properties = ( align => 'justify' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="justify"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 18. Test the _write_xf() method. Horizontal alignment = center across.
#
%properties = ( align => 'center_across' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="centerContinuous"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 19. Test the _write_xf() method. Horizontal alignment = distributed.
#
%properties = ( align => 'distributed' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 20. Test the _write_xf() method. Horizontal alignment = distributed + indent.
#
%properties = ( align => 'distributed', indent => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed" indent="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 21. Test the _write_xf() method. Horizontal alignment = justify distributed.
#
%properties = ( align => 'justify_distributed' );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed" justifyLastLine="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 22. Test the _write_xf() method. Horizontal alignment = indent only.
#     This should default to left alignment.
#
%properties = ( indent => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" indent="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );



###############################################################################
#
# 23. Test the _write_xf() method. Horizontal alignment = distributed + indent.
#     The justify_distributed should drop back to plain distributed if there
#     is an indent.
%properties = ( align => 'justify_distributed', indent => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="distributed" indent="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 24. Test the _write_xf() method. Alignment = text wrap
#
%properties = ( text_wrap => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 25. Test the _write_xf() method. Alignment = shrink to fit
#
%properties = ( shrink => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment shrinkToFit="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 26. Test the _write_xf() method. Alignment = reading order
#
%properties = ( reading_order => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment readingOrder="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 27. Test the _write_xf() method. Alignment = reading order
#
%properties = ( reading_order => 2 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment readingOrder="2"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 28. Test the _write_xf() method. Alignment = rotation
#
%properties = ( rotation => 45 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="45"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 29. Test the _write_xf() method. Alignment = rotation
#
%properties = ( rotation => -45 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="135"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 30. Test the _write_xf() method. Alignment = rotation
#
%properties = ( rotation => 270 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="255"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 31. Test the _write_xf() method. Alignment = rotation
#
%properties = ( rotation => 90 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="90"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 32. Test the _write_xf() method. Alignment = rotation
#
%properties = ( rotation => -90 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="180"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 33. Test the _write_xf() method. With cell protection.
#
%properties = ( locked => 0 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 34. Test the _write_xf() method. With cell protection.
#
%properties = ( hidden => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection hidden="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 35. Test the _write_xf() method. With cell protection.
#
%properties = ( locked => 0, hidden => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 36. Test the _write_xf() method. With cell protection + align.
#
%properties = ( align => 'right', locked => 0, hidden => 1 );

$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyProtection="1"><alignment horizontal="right"/><protection locked="0" hidden="1"/></xf>';

$format = Excel::Writer::XLSX::Format->new( {}, {}, %properties );

$style = _new_style(\$got);
$style->_write_xf( $format );

is( $got, $expected, $caption );


__END__


