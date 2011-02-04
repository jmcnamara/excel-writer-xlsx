###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# reverse('©'), February 2011, John McNamara, jmcnamara@cpan.org
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
# Test the _write_hyperlink() method.
#
$caption  = " \tWorksheet: _write_hyperlink()";
$expected = '<hyperlink ref="A1" r:id="rId1" />';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_hyperlink( 0, 0, 1 );

is( $got, $expected, $caption );

__END__


