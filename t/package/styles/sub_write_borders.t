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
# Test the _write_borders() method.
#
$caption  = " \tStyles: _write_borders()";
$expected = '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>';

my @formats = ( Excel::Writer::XLSX::Format->new( 0, {}, has_border => 1 ) );

$style = _new_style(\$got);
$style->_set_style_properties( \@formats, undef, undef, undef, 1 );
$style->_write_borders();

is( $got, $expected, $caption );

__END__


