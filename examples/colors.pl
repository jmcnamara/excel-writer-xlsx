#!/usr/bin/perl -w

################################################################################
#
# Demonstrates Excel::Writer::XLSX's named colours and the Excel colour
# palette.
#
# The set_custom_color() Worksheet method can be used to override one of the
# built-in palette values with a more suitable colour. See the main docs.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

my $workbook = Excel::Writer::XLSX->new( 'colors.xlsx' );

# Some common formats
my $center = $workbook->add_format( align => 'center' );
my $heading = $workbook->add_format( align => 'center', bold => 1 );


######################################################################
#
# Demonstrate the named colors.
#

my %colors = (
    0x08, 'black',
    0x0C, 'blue',
    0x10, 'brown',
    0x0F, 'cyan',
    0x17, 'gray',
    0x11, 'green',
    0x0B, 'lime',
    0x0E, 'magenta',
    0x12, 'navy',
    0x35, 'orange',
    0x21, 'pink',
    0x14, 'purple',
    0x0A, 'red',
    0x16, 'silver',
    0x09, 'white',
    0x0D, 'yellow',

);

my $worksheet1 = $workbook->add_worksheet( 'Named colors' );

$worksheet1->set_column( 0, 3, 15 );

$worksheet1->write( 0, 0, "Index", $heading );
$worksheet1->write( 0, 1, "Index", $heading );
$worksheet1->write( 0, 2, "Name",  $heading );
$worksheet1->write( 0, 3, "Color", $heading );

my $i = 1;

while ( my ( $index, $color ) = each %colors ) {
    my $format = $workbook->add_format(
        fg_color => $color,
        pattern  => 1,
        border   => 1
    );

    $worksheet1->write( $i + 1, 0, $index, $center );
    $worksheet1->write( $i + 1, 1, sprintf( "0x%02X", $index ), $center );
    $worksheet1->write( $i + 1, 2, $color, $center );
    $worksheet1->write( $i + 1, 3, '',     $format );
    $i++;
}


######################################################################
#
# Demonstrate the standard Excel colors in the range 8..63.
#

my $worksheet2 = $workbook->add_worksheet( 'Standard colors' );

$worksheet2->set_column( 0, 3, 15 );

$worksheet2->write( 0, 0, "Index", $heading );
$worksheet2->write( 0, 1, "Index", $heading );
$worksheet2->write( 0, 2, "Color", $heading );
$worksheet2->write( 0, 3, "Name",  $heading );

for my $i ( 8 .. 63 ) {
    my $format = $workbook->add_format(
        fg_color => $i,
        pattern  => 1,
        border   => 1
    );

    $worksheet2->write( ( $i - 7 ), 0, $i, $center );
    $worksheet2->write( ( $i - 7 ), 1, sprintf( "0x%02X", $i ), $center );
    $worksheet2->write( ( $i - 7 ), 2, '', $format );

    # Add the  color names
    if ( exists $colors{$i} ) {
        $worksheet2->write( ( $i - 7 ), 3, $colors{$i}, $center );

    }
}


######################################################################
#
# Demonstrate the Html colors.
#



%colors = (
	'#000000',  'black',
	'#0000FF',  'blue',
	'#800000',  'brown',
	'#00FFFF',  'cyan',
	'#808080',  'gray',
	'#008000',  'green',
	'#00FF00',  'lime',
	'#FF00FF',  'magenta',
	'#000080',  'navy',
	'#FF6600',  'orange',
	'#FF00FF',  'pink',
	'#800080',  'purple',
	'#FF0000',  'red',
	'#C0C0C0',  'silver',
	'#FFFFFF',  'white',
	'#FFFF00',  'yellow',
);

my $worksheet3 = $workbook->add_worksheet( 'Html colors' );

$worksheet3->set_column( 0, 3, 15 );

$worksheet3->write( 0, 0, "Html", $heading );
$worksheet3->write( 0, 1, "Name",  $heading );
$worksheet3->write( 0, 2, "Color", $heading );

$i = 1;

while ( my ( $html_color, $color ) = each %colors ) {
    my $format = $workbook->add_format(
        fg_color => $html_color,
        pattern  => 1,
        border   => 1
    );

    $worksheet3->write( $i + 1, 1, $html_color, $center );
    $worksheet3->write( $i + 1, 2, $color,      $center );
    $worksheet3->write( $i + 1, 3, '',          $format );
    $i++;
}

$workbook->close();

__END__
