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
# Test the _write_mru_colors() method.
#
$caption  = " \tStyles: _write_mru_colors()";
$expected = '<mruColors><color rgb="FF26DA55"/></mruColors>';

$style = _new_style(\$got);

$style->_write_mru_colors( 'FF26DA55' );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_mru_colors() method.
#
$caption  = " \tStyles: _write_mru_colors()";
$expected = '<mruColors><color rgb="FF646462"/><color rgb="FF792DC8"/><color rgb="FF26DA55"/></mruColors>';

$style = _new_style(\$got);

$style->_write_mru_colors( 'FF26DA55', 'FF792DC8', 'FF646462' );

is( $got, $expected, $caption );

__END__


