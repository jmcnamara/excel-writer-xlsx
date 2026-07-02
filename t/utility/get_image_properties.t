###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2025, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use warnings;
use File::Temp qw(tempfile);
use Excel::Writer::XLSX::Utility qw(get_image_properties);

use Test::More tests => 25;

###############################################################################
#
# Tests setup.
#
my $dir = 't/regression/images/';

my ( $type, $width, $height, $name, $x_dpi, $y_dpi, $md5 );


###############################################################################
#
# Test the get_image_properties() function.
#
my @tests = (

    # filename                  type   width  height  x_dpi  y_dpi
    [ 'red.png',               'png',    32,     32,    96,    96 ],
    [ 'red_64x20.png',         'png',    64,     20,    96,    96 ],
    [ 'red_topdown_32x32.bmp', 'bmp',    32,     32,    96,    96 ],
    [ 'red_topdown_20x12.bmp', 'bmp',    20,     12,    96,    96 ],
    [ 'black_40000x45000.gif', 'gif', 40000,  45000,    96,    96 ],

);


for my $test ( @tests ) {
    my ( $filename, $exp_type, $exp_width, $exp_height, $exp_x_dpi, $exp_y_dpi )
      = @{$test};

    ( $type, $width, $height, $name, $x_dpi, $y_dpi, $md5 ) =
      get_image_properties( $dir . $filename );

    my $caption = " \tUtility: get_image_properties( $filename ) ->";

    is( $type,   $exp_type,   "$caption type" );
    is( $width,  $exp_width,  "$caption width" );
    is( $height, $exp_height, "$caption height" );
    is( $x_dpi,  $exp_x_dpi,  "$caption x_dpi" );
    is( $y_dpi,  $exp_y_dpi,  "$caption y_dpi" );
}

__END__
