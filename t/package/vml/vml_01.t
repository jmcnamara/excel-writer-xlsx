###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::VML methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
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

$vml->_assemble_xml_file(
    1, 1024,
    [
        [
            1, 1, 'Some text', '', undef, '#ffffe1', 'Tahoma', 8, 2,
            [ 2, 0, 15, 10, 4, 4, 15, 4, 143, 10, 128, 74 ]
        ]
    ]
);

$expected = _expected_vml_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"
  path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path gradientshapeok="t" o:connecttype="rect"/>
 </v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t202" style='position:absolute;
  margin-left:107.25pt;margin-top:7.5pt;width:96pt;height:55.5pt;z-index:1;
  visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto">
  <v:fill color2="#ffffe1"/>
  <v:shadow on="t" color="black" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style='mso-direction-alt:auto'>
   <div style='text-align:left'></div>
  </v:textbox>
  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>
    2, 15, 0, 10, 4, 15, 4, 4</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>1</x:Row>
   <x:Column>1</x:Column>
  </x:ClientData>
 </v:shape></xml>
