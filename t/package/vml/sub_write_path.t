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

use Test::More tests => 3;


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
# Test the _write_path() method.
#
$caption  = " \tVML: _write_path()";
$expected = '<v:path gradientshapeok="t" o:connecttype="rect"/>';

$vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );

$vml->_write_comment_path( 't', 'rect');

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_path() method.
#
$caption  = " \tVML: _write_path()";
$expected = '<v:path o:connecttype="none"/>';

$vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );

$vml->_write_comment_path( undef, 'none');

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_path() method.
#
$caption  = " \tVML: _write_path()";
$expected = '<v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>';

$vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );

$vml->_write_button_path();

is( $got, $expected, $caption );

__END__


