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
# Test the _write_hyperlinks() method.
#
$caption  = " \tWorksheet: _write_hyperlinks()";
$expected = '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>';

$worksheet = _new_worksheet(\$got);

$worksheet->{_hlink_refs} = [[ 1, 0, 0, 1 ]];
$worksheet->_write_hyperlinks();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_hyperlinks() method.
#
$caption  = " \tWorksheet: _write_hyperlinks()";
$expected = '<hyperlinks><hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1"/></hyperlinks>';

$worksheet = _new_worksheet(\$got);

$worksheet->{_hlink_refs} = [[ 2, 0, 0, 'Sheet2!A1', 'Sheet2!A1' ]];
$worksheet->_write_hyperlinks();

is( $got, $expected, $caption );

__END__


