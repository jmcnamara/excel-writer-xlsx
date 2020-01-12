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
# Test the _write_table_style_info() method.
#
$caption  = " \tTable: _write_table_style_info()";
$expected = '<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>';

$table = _new_object( \$got, 'Excel::Writer::XLSX::Package::Table' );

$table->{_properties} = {
    _style            => 'TableStyleMedium9',
    _show_first_col   => 0,
    _show_last_col    => 0,
    _show_row_stripes => 1,
    _show_col_stripes => 0,
};



$table->_write_table_style_info();

is( $got, $expected, $caption );

__END__


