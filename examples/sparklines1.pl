#!/usr/bin/perl

###############################################################################
#
# Example of how to add sparklines to an Excel::Writer::XLSX file.
#
# Sparklines are small charts that fit in a single cell and are
# used to show trends in data. See sparklines2.pl for examples
# of more complex sparkline formatting.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'sparklines1.xlsx' );
my $worksheet = $workbook->add_worksheet();

# Some sample data to plot.
my $data = [

    [ -2, 2,  3,  -1, 0 ],
    [ 30, 20, 33, 20, 15 ],
    [ 1,  -1, -1, 1,  -1 ],

];

# Write the sample data to the worksheet.
$worksheet->write_col( 'A1', $data );


# Add a line sparkline (the default) with markers.
$worksheet->add_sparkline(
    {
        location => 'F1',
        range    => 'Sheet1!A1:E1',
        markers  => 1,
    }
);

# Add a column sparkline with non-default style.
$worksheet->add_sparkline(
    {
        location => 'F2',
        range    => 'Sheet1!A2:E2',
        type     => 'column',
        style    => 12,
    }
);

# Add a win/loss sparkline with negative values highlighted.
$worksheet->add_sparkline(
    {
        location        => 'F3',
        range           => 'Sheet1!A3:E3',
        type            => 'win_loss',
        negative_points => 1,
    }
);

$workbook->close();

__END__
