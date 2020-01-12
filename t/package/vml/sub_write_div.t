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

use Test::More tests => 1;


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
# Test the _write_div() method.
#
$caption  = " \tVML: _write_div()";
$expected = '<div style="text-align:left"></div>';

$vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );

$vml->_write_div( 'left' );

is( $got, $expected, $caption );

__END__


