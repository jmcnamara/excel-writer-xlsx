###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_style _new_workbook);
use strict;
use warnings;

use Test::More tests => 17;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $style;
my $format;
my %properties;

# Generate a temporary Workbook object to get the colour palette.
my $tmp      = '';
my $workbook = _new_workbook(\$tmp);
my $palette  = $workbook->{_palette};


###############################################################################
#
# 1. Test the _write_font() method. Default properties.
#
%properties = ();
$caption    = " \tStyles: _write_font()";
$expected   = '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {} );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_font() method. Bold.
#
%properties = ( bold => 1 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_font() method. Italic.
#
%properties = ( italic => 1 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><i/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_font() method. Underline.
#
%properties = ( underline => 1 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><u/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_font() method. Strikeout.
#
%properties = ( font_strikeout => 1 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><strike/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_font() method. Superscript.
#
%properties = ( font_script => 1 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><vertAlign val="superscript"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 7. Test the _write_font() method. Subscript.
#
%properties = ( font_script => 2 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><vertAlign val="subscript"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 8. Test the _write_font() method. Font name.
#
%properties = ( font => 'Arial' );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><sz val="11"/><color theme="1"/><name val="Arial"/><family val="2"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 9. Test the _write_font() method. Font size.
#
%properties = ( size => 12 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 10. Test the _write_font() method. Outline.
#
%properties = ( font_outline => 1 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><outline/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 11. Test the _write_font() method. Shadow.
#
%properties = ( font_shadow => 1 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><shadow/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 12. Test the _write_font() method. Colour = red.
#
%properties = ( color => 'red' );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style = _new_style( \$got );
$style->_set_style_properties( undef, $palette, undef, undef );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 13. Test the _write_font() method. All font attributes to check order.
#
%properties = (
    bold           => 1,
    color          => 'red',
    font_outline   => 1,
    font_script    => 1,
    font_shadow    => 1,
    font_strikeout => 1,
    italic         => 1,
    size           => 12,
    underline      => 1,
);

$caption    = " \tStyles: _write_font()";
$expected   = '<font><b/><i/><strike/><outline/><shadow/><u/><vertAlign val="superscript"/><sz val="12"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style = _new_style( \$got );
$style->_set_style_properties( undef, $palette, undef, undef );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 14. Test the _write_font() method. Double underline.
#
%properties = ( underline => 2 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><u val="double"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 15. Test the _write_font() method. Double underline.
#
%properties = ( underline => 33 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><u val="singleAccounting"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 16. Test the _write_font() method. Double underline.
#
%properties = ( underline => 34 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><u val="doubleAccounting"/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );


###############################################################################
#
# 17. Test the _write_font() method. Hyperlink.
#
%properties = ( hyperlink => 1 );
$caption    = " \tStyles: _write_font()";
$expected   = '<font><u/><sz val="11"/><color theme="10"/><name val="Calibri"/><family val="2"/></font>';

$format = Excel::Writer::XLSX::Format->new( 0, {}, %properties );
$style  = _new_style( \$got );
$style->_write_font( $format );

is( $got, $expected, $caption );



__END__


