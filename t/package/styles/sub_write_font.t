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
# Test the _write_font() method.
#
$caption = " \tStyles: _write_font()";
$expected =
'<font><sz val="11" /><color theme="1" /><name val="Calibri" /><family val="2" /><scheme val="minor" /></font>';

my $format = Excel::Writer::XLSX::Format->new( 0 );

$style = _new_style(\$got);

$style->_write_font( $format );

is( $got, $expected, $caption );

__END__


