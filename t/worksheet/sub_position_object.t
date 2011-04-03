###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Tests for the _position_object method used to calcualte the twoCellAnchor
# positions for drawing and chart objects.
#
# reverse('©'), March 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_new_worksheet _is_deep_diff);
use strict;
use warnings;

use Test::More tests => 3;

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
# 1. Test _position_object() for chart vertices.
#
$caption = " \tWorksheet: _position_object()";
@expected = ( 4, 8, 0, 0, 11, 22, 304800, 76200 );

$worksheet = _new_worksheet( \$tmp );

@got = $worksheet->_position_object( 4, 8, 0, 0, 480, 288 );

_is_deep_diff( \@got, \@expected, $caption );


###############################################################################
#
# 2. Test _position_object() for chart vertices.
#
$caption = " \tWorksheet: _position_object()";
@expected = ( 4, 8, 0, 0, 11, 22, 0, 76200 );

$worksheet = _new_worksheet( \$tmp );
$worksheet->set_column( 'L:L', 3.86 );

@got = $worksheet->_position_object( 4, 8, 0, 0, 480, 288 );

_is_deep_diff( \@got, \@expected, $caption );


###############################################################################
#
# 3. Test _position_object() for chart vertices.
#
$caption = " \tWorksheet: _position_object()";
@expected = ( 4, 8, 0, 0, 11, 22, 0, 0 );

$worksheet = _new_worksheet( \$tmp );
$worksheet->set_column( 'L:L', 3.86 );
$worksheet->set_row( 22, 6 );

@got = $worksheet->_position_object( 4, 8, 0, 0, 480, 288 );

_is_deep_diff( \@got, \@expected, $caption );


__END__
