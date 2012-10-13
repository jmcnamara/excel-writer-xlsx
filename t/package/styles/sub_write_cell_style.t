###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# reverse ('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_style';
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
my $style;


###############################################################################
#
# Test the _write_cell_style() method.
#
$caption  = " \tStyles: _write_cell_style()";
$expected = '<cellStyle name="Normal" xfId="0" builtinId="0"/>';

$style = _new_style(\$got);

$style->_write_cell_style();

is( $got, $expected, $caption );

__END__


