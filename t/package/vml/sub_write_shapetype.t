###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::VML methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_object';
use strict;
use warnings;
use Excel::Writer::XLSX::Package::VML;

use Test::More tests => 2;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $vml;


###############################################################################
#
# Test the _write_shapetype() method.
#
$caption  = " \tVML: _write_shapetype()";
$expected = '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';


$vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );

$vml->_write_comment_shapetype();

is( $got, $expected, $caption );



###############################################################################
#
# Test the _write_shapetype() method.
#
$caption  = " \tVML: _write_shapetype()";
$expected = '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>';


$vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );

$vml->_write_button_shapetype();

is( $got, $expected, $caption );

__END__


