###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# reverse('�'), September 2010, John McNamara, jmcnamara@cpan.org
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
# Test the _write_xf() method.
#
$caption  = " \tStyles: _write_xf()";
$expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />';

my $format = Excel::Writer::XLSX::Format->new( 0 );

$style = _new_style(\$got);

$style->_write_xf( $format );

is( $got, $expected, $caption );

__END__


