###############################################################################
#
# Tests for Excel::Writer::XLSX::Chart methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_object);
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
my $tmp = '';
my $caption;
my $chart;


###############################################################################
#
# Test the _process_names() method.
#
$caption  = " \tChart: _process_names()";
$expected = [ undef, undef ];

$chart = _new_object( \$tmp, 'Excel::Writer::XLSX::Chart' );

$got = [ $chart->_process_names() ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# Test the _process_names() method.
#
$caption  = " \tChart: _process_names()";
$expected = [ 'Text', undef ];

$chart = _new_object( \$tmp, 'Excel::Writer::XLSX::Chart' );

$got = [ $chart->_process_names( 'Text' ) ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# Test the _process_names() method.
#
$caption  = " \tChart: _process_names()";
$expected = [ '', '=Sheet1!$A$1' ];

$chart = _new_object( \$tmp, 'Excel::Writer::XLSX::Chart' );

$got = [ $chart->_process_names( '=Sheet1!$A$1' ) ];

is_deeply( $got, $expected, $caption );



__END__


