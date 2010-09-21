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
# Test the _write_page_setup() method.
#
$caption  = " \tWorksheet: _write_page_setup()";
$expected = '<pageSetup paperSize="0" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292" />';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_page_setup();

is( $got, $expected, $caption );

__END__


