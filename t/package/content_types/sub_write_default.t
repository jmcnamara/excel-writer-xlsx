###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::ContentTypes methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Package::ContentTypes;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Package::ContentTypes' );


###############################################################################
#
# Test the _write_default() method.
#
$caption  = " \tContentTypes: _write_default()";
$expected = '<Default Extension="xml" ContentType="application/xml"/>';

$obj->_write_default( 'xml', 'application/xml' );

is( $got, $expected, $caption );

__END__


