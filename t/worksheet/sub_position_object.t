###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Tests for the _position_object_emus method used to calcualte the twoCellAnchor
# positions for drawing and chart objects.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_worksheet _is_deep_diff);
use strict;
use warnings;

use Test::More tests => 6;

###############################################################################
#
# Tests setup.
#
my @expected;
my @got;
my $tmp = '';
my $caption;
my $worksheet;


###############################################################################
#
# 1. Test _position_object_emus() for chart vertices.
#
$caption = " \tWorksheet: _position_object_emus()";
@expected = ( 4, 8, 0, 0, 11, 22, 304800, 76200, 2438400, 1524000 );

$worksheet = _new_worksheet( \$tmp );

@got = $worksheet->_position_object_emus( 4, 8, 0, 0, 480, 288 );

_is_deep_diff( \@got, \@expected, $caption );


###############################################################################
#
# 2. Test _position_object_emus() for chart vertices.
#
$caption = " \tWorksheet: _position_object_emus()";
@expected = ( 4, 8, 0, 0, 12, 22, 0, 76200, 2438400, 1524000 );

$worksheet = _new_worksheet( \$tmp );
$worksheet->set_column( 'L:L', 3.86 );

@got = $worksheet->_position_object_emus( 4, 8, 0, 0, 480, 288 );

_is_deep_diff( \@got, \@expected, $caption );


###############################################################################
#
# 3. Test _position_object_emus() for chart vertices.
#
$caption = " \tWorksheet: _position_object_emus()";
@expected = ( 4, 8, 0, 0, 12, 23, 0, 0, 2438400, 1524000 );

$worksheet = _new_worksheet( \$tmp );
$worksheet->set_column( 'L:L', 3.86 );
$worksheet->set_row( 22, 6 );

@got = $worksheet->_position_object_emus( 4, 8, 0, 0, 480, 288 );

_is_deep_diff( \@got, \@expected, $caption );


###############################################################################
#
# 4. Test _position_object_emus() for image vertices.
#
$caption = " \tWorksheet: _position_object_emus()";
@expected = ( 4, 8, 0, 0, 4, 9, 304800, 114300, 2438400, 1524000 );

$worksheet = _new_worksheet( \$tmp );

@got = $worksheet->_position_object_emus( 4, 8, 0, 0, 32, 32 );

_is_deep_diff( \@got, \@expected, $caption );


###############################################################################
#
# 5. Test _position_object_emus() for image vertices.
#
$caption = " \tWorksheet: _position_object_emus()";
@expected = ( 4, 8, 19050, 28575, 5, 11, 95250, 142875, 2457450, 1552575 );

$worksheet = _new_worksheet( \$tmp );

@got = $worksheet->_position_object_emus( 4, 8, 2, 3, 72, 72 );

_is_deep_diff( \@got, \@expected, $caption );


###############################################################################
#
# 6. Test _position_object_emus() for image vertices.
#
$caption = " \tWorksheet: _position_object_emus()";
@expected = ( 5, 1, 19050, 28575, 6, 4, 352425, 114300, 3067050, 219075 );

$worksheet = _new_worksheet( \$tmp );

@got = $worksheet->_position_object_emus( 5, 1, 2, 3, 99, 69 );

_is_deep_diff( \@got, \@expected, $caption );


__END__
