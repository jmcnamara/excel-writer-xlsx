###############################################################################
#
# Tests for Excel::Writer::XLSX::Drawing methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Worksheet;
use Excel::Writer::XLSX::Shape;

use Test::More tests => 2;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;

our $WARN_TEXT;
$SIG{__WARN__} = sub { $WARN_TEXT = shift; };

my $sheet     = Excel::Writer::XLSX::Worksheet->new();
my $shape     = Excel::Writer::XLSX::Shape->new();
my $inserted1 = $sheet->insert_shape( 4, 8, $shape, 300, 400 );
my $inserted2 = $sheet->insert_shape( 8, 12, $shape, 500, 750 );

my $cxn_shape = Excel::Writer::XLSX::Shape->new(
    undef,
    name => 'link',
    type => 'bentConnector3'
);

###############################################################################
#
# Test missing start connection
#

$cxn_shape->set_start( 9999 );    # bogus shape id
$cxn_shape->set_start_index( 4 );
$cxn_shape->set_start_side( 'b' );

$cxn_shape->set_end( $inserted2->get_id() );
$cxn_shape->set_end_index( 0 );     # 0 - top connection point
$cxn_shape->set_end_side( 't' );    # l)eft or t)op

$sheet->insert_shape( 1, 1, $cxn_shape );
$caption  = " \tWorksheet: _auto_locate_connector() - missing start connection";
$expected = "missing start connection for 'link', id=9999\n";
_is_deep_diff( \$WARN_TEXT, \$expected, $caption );

###############################################################################
#
# Test missing end connection
#

$caption = " \tWorksheet: _auto_locate_connector() - missing end connection";
$cxn_shape->set_start( $inserted1->get_id() );
$cxn_shape->set_end( 9999 );    # bogus shape id
$sheet->insert_shape( 1, 1, $cxn_shape );
$expected = "missing end connection for 'link', id=9999\n";
_is_deep_diff( \$WARN_TEXT, \$expected, $caption );

