#!/usr/bin/perl

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# add shapes and one-to-many connectors to an Excel xlsx file.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'shape7.xlsx' );
my $worksheet = $workbook->add_worksheet();

# Add a circle, with centered text. c is for circle, not center.
my $cw = 60;
my $ch = 60;
my $cx = 210;
my $cy = 190;

my $ellipse = $workbook->add_shape(
    type   => 'ellipse',
    id     => 2,
    text   => "Hello\nWorld",
    width  => $cw,
    height => $ch
);
$worksheet->insert_shape( 'A1', $ellipse, $cx, $cy );

# Add a plus sign at 4 different positions around the circle.
my $pw = 20;
my $ph = 20;
my $px = 120;
my $py = 250;
my $plus =
  $workbook->add_shape( type => 'plus', id => 3, width => $pw, height => $ph );
my $p1 = $worksheet->insert_shape( 'A1', $plus, 350, 350 );
my $p2 = $worksheet->insert_shape( 'A1', $plus, 150, 350 );
my $p3 = $worksheet->insert_shape( 'A1', $plus, 350, 150 );
$plus->set_adjustments( 35 );    # change shape of plus symbol.
my $p4 = $worksheet->insert_shape( 'A1', $plus, 150, 150 );

my $cxn_shape = $workbook->add_shape( type => 'bentConnector3', fill => 0 );

$cxn_shape->set_start( $ellipse->get_id() );
$cxn_shape->set_start_index( 4 );    # 4nd connection pt, clockwise from top(0).
$cxn_shape->set_start_side( 'b' );   # r)ight or b)ottom.

$cxn_shape->set_end( $p1->get_id() );
$cxn_shape->set_end_index( 0 );
$cxn_shape->set_end_side( 't' );
$worksheet->insert_shape( 'A1', $cxn_shape, 0, 0 );

$cxn_shape->set_end( $p2->get_id() );
$worksheet->insert_shape( 'A1', $cxn_shape, 0, 0 );

$cxn_shape->set_end( $p3->get_id() );
$worksheet->insert_shape( 'A1', $cxn_shape, 0, 0 );

$cxn_shape->set_end( $p4->get_id() );
$cxn_shape->set_adjustments( -50, 45, 120 );
$worksheet->insert_shape( 'A1', $cxn_shape, 0, 0 );

$workbook->close();

__END__

