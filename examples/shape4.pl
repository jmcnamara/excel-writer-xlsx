#!/usr/bin/perl

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# demonstrate stenciling in an Excel xlsx file.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'shape4.xlsx' );
my $worksheet = $workbook->add_worksheet();

$worksheet->hide_gridlines( 2 );

my $type  = 'rect';
my $shape = $workbook->add_shape(
    type   => $type,
    width  => 90,
    height => 90,
);

for my $n ( 1 .. 10 ) {

    # Change the last 5 rectangles to stars. Previously inserted shapes stay
    # as rectangles.
    $type = 'star5' if $n == 6;
    $shape->set_type( $type );
    $shape->set_text( "$type $n" );
    $worksheet->insert_shape( 'A1', $shape, $n * 100, 50 );
}


my $stencil = $workbook->add_shape(
    stencil => 1,                    # The default.
    width   => 90,
    height  => 90,
    text    => 'started as a box',
);
$worksheet->insert_shape( 'A1', $stencil, 100, 150 );

$stencil->set_stencil( 0 );
$worksheet->insert_shape( 'A1', $stencil, 200, 150 );
$worksheet->insert_shape( 'A1', $stencil, 300, 150 );

# Ooops!  Changed my mind.  Change the rectangle to an ellipse (circle),
# for the last two shapes.
$stencil->set_type( 'ellipse' );
$stencil->set_text( 'Now its a circle' );

$workbook->close();

__END__
