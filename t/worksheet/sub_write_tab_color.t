###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
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
my $worksheet;


###############################################################################
#
# Test the _write_tab_color() method.
#
$caption  = " \tWorksheet: _write_tab_color()";
$expected = '<tabColor rgb="FFFF0000"/>';

$worksheet = _new_worksheet(\$got);
# Mock up the color palette.
$worksheet->{_tab_color} = 0x0A;
$worksheet->{_palette}->[2] = [ 0xff, 0x00, 0x00, 0x00 ];


$worksheet->_write_tab_color(  );

is( $got, $expected, $caption );

__END__


