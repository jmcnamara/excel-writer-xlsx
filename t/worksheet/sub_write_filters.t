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

use Test::More tests => 3;


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
# Test the _write_filters() method.
#
$caption  = " \tWorksheet: _write_filters()";
$expected = '<filters><filter val="East"/></filters>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_filters( 'East' );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_filters() method.
#
$caption  = " \tWorksheet: _write_filters()";
$expected = '<filters><filter val="East"/><filter val="South"/></filters>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_filters( 'East', 'South' );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_filters() method.
#
$caption  = " \tWorksheet: _write_filters()";
$expected = '<filters blank="1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_filters( 'blanks' );

is( $got, $expected, $caption );

__END__


