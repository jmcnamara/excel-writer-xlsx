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

use Test::More tests => 1;


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
# Test the _write_selection() method.
#
$caption  = " \tWorksheet: _write_selection()";
$expected = '<selection activeCell="A1" sqref="A1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_selection( undef, 'A1', 'A1' );

is( $got, $expected, $caption );

__END__


