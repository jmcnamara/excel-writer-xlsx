###############################################################################
#
# Tests for Excel::Writer::XLSX::Chart methods.
#
# reverse('(c)'), March 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_object';
use strict;
use warnings;
use Excel::Writer::XLSX::Chart;

use Test::More tests => 1;


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
# Test the _write_legend_pos() method.
#
$caption  = " \tChart: _write_legend_pos()";
$expected = '<c:legendPos val="r"/>';

$chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );

$chart->_write_legend_pos( 'r' );

is( $got, $expected, $caption );

__END__


