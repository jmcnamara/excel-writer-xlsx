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
# Test the _write_table_column() method.
#
$caption  = " \tTable: _write_table_column()";
$expected = '<tableColumn id="1" name="Column1"/>';

$table = _new_object( \$got, 'Excel::Writer::XLSX::Package::Table' );

$table->_write_table_column( { _name => 'Column1', _id => 1 } );

is( $got, $expected, $caption );

__END__


