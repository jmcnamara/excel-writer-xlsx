###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use warnings;
use Excel::Writer::XLSX::Worksheet;

use Test::More tests => 2337;


###############################################################################
#
# Tests setup.
#


# Function for testing.
sub width_to_pixels {

    my $width           = shift;
    my $max_digit_width = 7;
    my $padding         = 5;
    my $pixels;

    if ( $width < 1 ) {
        $pixels = int( $width * ( $max_digit_width + $padding ) + 0.5 );
    }
    else {
        $pixels = int( $width * $max_digit_width + 0.5 ) + $padding;
    }

    return $pixels;
}

# Function for testing.
sub height_to_pixels {

    my $height = shift;

    return int( 4 / 3 * $height );
}


###############################################################################
#
# Test _pixel_to_width().
#
for my $pixels ( 0 .. 1790 ) {

    my $caption  = " \tWorksheet: _pixel_to_width()";
    my $got      = width_to_pixels( Excel::Writer::XLSX::Worksheet::_pixels_to_width( $pixels ) );
    my $expected = $pixels;

    is( $got, $expected, $caption );

}


###############################################################################
#
# Test _pixel_to_height().
#
for my $pixels ( 0 .. 545 ) {

    my $caption  = " \tWorksheet: _pixel_to_height()";
    my $got      = height_to_pixels( Excel::Writer::XLSX::Worksheet::_pixels_to_height( $pixels ) );
    my $expected = $pixels;

    is( $got, $expected, $caption );

}


__END__
