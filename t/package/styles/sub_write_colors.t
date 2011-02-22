###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# reverse('(c)'), February 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_style';
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
my $style;


###############################################################################
#
# Test the _write_colors() method.
#
$caption  = " \tStyles: _write_colors()";
$expected = '<colors><mruColors><color rgb="FF26DA55" /></mruColors></colors>';

$style = _new_style(\$got);

$style->{_custom_colors} = [ 'FF26DA55' ];

$style->_write_colors();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_colors() method.
#
$caption  = " \tStyles: _write_colors()";
$expected = '<colors><mruColors><color rgb="FF646462" /><color rgb="FF792DC8" /><color rgb="FF26DA55" /></mruColors></colors>';

$style = _new_style(\$got);

$style->{_custom_colors} = [ 'FF26DA55', 'FF792DC8', 'FF646462' ];

$style->_write_colors();

is( $got, $expected, $caption );


__END__


