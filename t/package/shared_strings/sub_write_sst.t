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
# Test the _write_sst() method.
#
$caption  = " \tSharedStrings: _write_sst()";
$expected = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">';

$obj->_write_sst( 7, 3 );

is( $got, $expected, $caption );

__END__


