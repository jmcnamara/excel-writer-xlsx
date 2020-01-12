###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 2;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;


###############################################################################
#
# Test the _write_custom_filters() method.
#
$caption  = " \tWorksheet: _write_custom_filters()";
$expected = '<customFilters><customFilter operator="greaterThan" val="4000"/></customFilters>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_custom_filters( 4, 4000 );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_custom_filters() method.
#
$caption  = " \tWorksheet: _write_custom_filters()";
$expected = '<customFilters and="1"><customFilter operator="greaterThan" val="3000"/><customFilter operator="lessThan" val="8000"/></customFilters>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_custom_filters( 4, 3000, 0, 1, 8000 );

is( $got, $expected, $caption );

__END__


