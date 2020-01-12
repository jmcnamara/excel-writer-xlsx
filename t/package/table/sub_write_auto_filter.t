###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Table methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_object';
use strict;
use warnings;
use Excel::Writer::XLSX::Package::Table;

use Test::More tests => 1;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $table;


###############################################################################
#
# Test the _write_auto_filter() method.
#
$caption  = " \tTable: _write_auto_filter()";
$expected = '<autoFilter ref="C3:F13"/>';

$table = _new_object( \$got, 'Excel::Writer::XLSX::Package::Table' );

$table->{_properties}->{_autofilter} = 'C3:F13';

$table->_write_auto_filter( 'C3:F13' );

is( $got, $expected, $caption );

__END__


