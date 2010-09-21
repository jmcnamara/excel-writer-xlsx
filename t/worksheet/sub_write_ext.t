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
# Test the _write_ext() method.
#
$caption  = " \tWorksheet: _write_ext()";
$expected = '<ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="1" OnePage="0" WScale="0" /></ext>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_ext();

is( $got, $expected, $caption );

__END__


