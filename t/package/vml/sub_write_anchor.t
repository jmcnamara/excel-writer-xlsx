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
# Test the _write_anchor() method.
#
$caption  = " \tVML: _write_anchor()";
$expected = '<x:Anchor>2, 15, 0, 10, 4, 15, 4, 4</x:Anchor>';

$vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );

$vml->_write_anchor( [ 2, 0, 15, 10, 4, 4, 15, 4 ] );

is( $got, $expected, $caption );

__END__


