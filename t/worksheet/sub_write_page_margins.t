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
# Test the _write_page_margins() method.
#
$caption  = " \tWorksheet: _write_page_margins()";
$expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_page_margins();

is( $got, $expected, $caption );

__END__


