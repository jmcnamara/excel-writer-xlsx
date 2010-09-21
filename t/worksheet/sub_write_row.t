###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# reverse('�'), September 2010, John McNamara, jmcnamara@cpan.org
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
# Test the _write_row() method.
#
$caption  = " \tWorksheet: _write_row()";
$expected = '<row r="1">';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_row( 0 );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_row() method.
#
$caption  = " \tWorksheet: _write_row()";
$expected = '<row r="3" spans="2:2">';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_row( 2, '2:2' );

is( $got, $expected, $caption );

__END__


