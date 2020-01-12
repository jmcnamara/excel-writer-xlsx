###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_style';
use strict;
use warnings;

use Test::More tests => 3;


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
$expected = '<colors><mruColors><color rgb="FF26DA55"/></mruColors></colors>';

$style = _new_style(\$got);

$style->{_custom_colors} = [ 'FF26DA55' ];

$style->_write_colors();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_colors() method.
#
$caption  = " \tStyles: _write_colors()";
$expected = '<colors><mruColors><color rgb="FF646462"/><color rgb="FF792DC8"/><color rgb="FF26DA55"/></mruColors></colors>';

$style = _new_style(\$got);

$style->{_custom_colors} = [ 'FF26DA55', 'FF792DC8', 'FF646462' ];

$style->_write_colors();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_colors() method. Test that the mruColors are limited to 10.
#
$caption  = " \tStyles: _write_colors()";
$expected = '<colors><mruColors><color rgb="FFD97827"/><color rgb="FFB97847"/><color rgb="FF913AC6"/><color rgb="FFE3FA06"/><color rgb="FF0CF49C"/><color rgb="FF600FF1"/><color rgb="FFA1A759"/><color rgb="FFE31DAF"/><color rgb="FF583AC6"/><color rgb="FF5EA29C"/></mruColors></colors>';

$style = _new_style(\$got);

$style->{_custom_colors} = [
    'FF792DC8', 'FF646462', 'FF5EA29C', 'FF583AC6', 'FFE31DAF', 'FFA1A759',
    'FF600FF1', 'FF0CF49C', 'FFE3FA06', 'FF913AC6', 'FFB97847', 'FFD97827'
];

$style->_write_colors();

is( $got, $expected, $caption );


__END__


