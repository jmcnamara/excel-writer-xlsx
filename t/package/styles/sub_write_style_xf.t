###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# reverse('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
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
# Test the _write_style_xf() method.
#
$caption  = " \tStyles: _write_style_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>';

$style = _new_style(\$got);

$style->_write_style_xf();

is( $got, $expected, $caption );

__END__


