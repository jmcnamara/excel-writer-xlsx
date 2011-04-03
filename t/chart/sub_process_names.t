###############################################################################
#
# Tests for Excel::Writer::XLSX::Chart methods.
#
# reverse('(c)'), March 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_object _is_deep_diff);
use strict;
use warnings;
use Excel::Writer::XLSX::Chart;

use Test::More tests => 3;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $chart;


###############################################################################
#
# Test the _process_names() method.
#
$caption  = " \tChart: _process_names()";
$expected = [ undef, undef ];

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$got = [ $chart->_process_names() ];

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test the _process_names() method.
#
$caption  = " \tChart: _process_names()";
$expected = [ 'Text', undef ];

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$got = [ $chart->_process_names( 'Text' ) ];

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test the _process_names() method.
#
$caption  = " \tChart: _process_names()";
$expected = [ '', '=Sheet1!$A$1' ];

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$got = [ $chart->_process_names( '=Sheet1!$A$1' ) ];

_is_deep_diff( $got, $expected, $caption );



__END__


