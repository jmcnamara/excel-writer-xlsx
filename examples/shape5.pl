#!/usr/bin/perl

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# add shapes (objects and top/bottom connectors) to an Excel xlsx file.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'shape5.xlsx' );
my $worksheet = $workbook->add_worksheet();

my $s1 = $workbook->add_shape( type => 'ellipse', width => 60, height => 60 );
$worksheet->insert_shape( 'A1', $s1, 50, 50 );

my $s2 = $workbook->add_shape( type => 'plus', width => 20, height => 20 );
$worksheet->insert_shape( 'A1', $s2, 250, 200 );

# Create a connector to link the two shapes.
my $cxn_shape = $workbook->add_shape( type => 'bentConnector3' );

# Link the start of the connector to the right side.
$cxn_shape->set_start( $s1->get_id() );
$cxn_shape->set_start_index( 4 );  # 4th connection pt, clockwise from top(0).
$cxn_shape->set_start_side( 'b' ); # r)ight or b)ottom.

# Link the end of the connector to the left side.
$cxn_shape->set_end( $s2->get_id() );
$cxn_shape->set_end_index( 0 );     # clockwise from top(0).
$cxn_shape->set_end_side( 't' );    # t)top.

$worksheet->insert_shape( 'A1', $cxn_shape, 0, 0 );

$workbook->close();

__END__
