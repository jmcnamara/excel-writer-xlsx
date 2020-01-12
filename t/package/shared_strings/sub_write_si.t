###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::SharedStrings methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::SharedStrings;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Package::SharedStrings' );


###############################################################################
#
# Test the _write_si() method.
#
$caption  = " \tSharedStrings: _write_si()";
$expected = '<si><t>neptune</t></si>';

$obj->_write_si('neptune');

is( $got, $expected, $caption );

__END__


