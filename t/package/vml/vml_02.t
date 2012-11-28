###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::VML methods.
#
# reverse ('(c)'), September 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_vml_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::VML;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );


###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tVML: _assemble_xml_file()";

$vml->_assemble_xml_file( 1, 1024, undef, [ [ 1, 1, 'Some text', '', undef, '#ffffe1', [ 2, 0, 15, 10, 4, 4, 15, 4, 143, 10, 128, 74 ] ] ] );

$expected = _expected_vml_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1"/>
  </o:shapelayout>
  <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">
    <v:stroke joinstyle="miter"/>
    <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
    <o:lock v:ext="edit" shapetype="t"/>
  </v:shapetype>
  <v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute;margin-left:96pt;margin-top:15pt;width:48pt;height:15pt;z-index:1;mso-wrap-style:tight" o:button="t" fillcolor="buttonFace [67]" strokecolor="windowText [64]" o:insetmode="auto">
    <v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>
    <o:lock v:ext="edit" rotation="t"/>
    <v:textbox style="mso-direction-alt:auto" o:singleclick="f">
      <div style="text-align:center">
        <font face="Calibri" size="220" color="#000000">Button 1</font>
      </div>
    </v:textbox>
    <x:ClientData ObjectType="Button">
      <x:Anchor>2, 0, 1, 0, 3, 0, 2, 0</x:Anchor>
      <x:PrintObject>False</x:PrintObject>
      <x:AutoFill>False</x:AutoFill>
      <x:FmlaMacro>[0]!Button1_Click</x:FmlaMacro>
      <x:TextHAlign>Center</x:TextHAlign>
      <x:TextVAlign>Center</x:TextVAlign>
    </x:ClientData>
  </v:shape>
</xml>
