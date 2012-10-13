###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::VML methods.
#
# reverse ('(c)'), September 2011, John McNamara, jmcnamara@cpan.org
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
# Test the _write_fill() method.
#
$caption  = " \tVML: _write_fill()";
$expected = '<v:fill color2="#ffffe1"/>';

$vml = _new_object( \$got, 'Excel::Writer::XLSX::Package::VML' );

$vml->_write_fill();

is( $got, $expected, $caption );

__END__


