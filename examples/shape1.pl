#!/usr/bin/perl

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# add shapes to an Excel xlsx file.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'shape1.xlsx' );
my $worksheet = $workbook->add_worksheet();

# Add a circle, with centered text.
my $ellipse = $workbook->add_shape(
    type   => 'ellipse',
    text   => "Hello\nWorld",
    width  => 60,
    height => 60
);

$worksheet->insert_shape( 'A1', $ellipse, 50, 50 );

# Add a plus sign.
my $plus = $workbook->add_shape( type => 'plus', width => 20, height => 20 );
$worksheet->insert_shape( 'D8', $plus );

$workbook->close();

__END__
