#!/usr/bin/perl

###############################################################################
#
# Example of how to add conditional formatting to an Excel::Writer::XLSX file.
#
# Conditional formatting allows you to apply a format to a cell or a range of
# cells based on a certain criteria.
#
# reverse('©'), October 2011, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook   = Excel::Writer::XLSX->new( 'conditional_format.xlsx' );
my $worksheet1 = $workbook->add_worksheet();
my $worksheet2 = $workbook->add_worksheet();
my $worksheet3 = $workbook->add_worksheet();



# Light red fill with dark red text.
my $format1 = $workbook->add_format(
    bg_color => '#FFC7CE',
    color    => '#9C0006',

);

# Green fill with dark green text.
my $format2 = $workbook->add_format(
    bg_color => '#C6EFCE',
    color    => '#006100',

);

# Some sample data to run the conditional formatting against.
my $data = [
    [ 90, 80,  50, 10,  20,  90,  40, 90,  30,  40 ],
    [ 20, 10,  90, 100, 30,  60,  70, 60,  50,  90 ],
    [ 10, 50,  60, 50,  20,  50,  80, 30,  40,  60 ],
    [ 10, 90,  20, 40,  10,  40,  50, 70,  90,  50 ],
    [ 70, 100, 10, 90,  10,  10,  20, 100, 100, 40 ],
    [ 20, 60,  10, 100, 30,  10,  20, 60,  100, 10 ],
    [ 10, 60,  10, 80,  100, 80,  30, 30,  70,  40 ],
    [ 30, 90,  60, 10,  10,  100, 40, 40,  30,  40 ],
    [ 80, 90,  10, 20,  20,  50,  80, 20,  60,  90 ],
    [ 60, 80,  30, 30,  10,  50,  80, 60,  50,  30 ],
];


###############################################################################
#
# Example 1.
#
my $caption = 'Cells with values >= 50 are in light red. '
  . 'Values < 50 are in light green.';

# Write the data.
$worksheet1->write( 'A1', $caption );
$worksheet1->write_col( 'B3', $data );

# Write a conditional format over a range.
$worksheet1->conditional_formatting( 'B3:K12',
    {
        type     => 'cell',
        format   => $format1,
        criteria => '>=',
        value    => 50,
    }
);

# Write another conditional format over the same range.
$worksheet1->conditional_formatting( 'B3:K12',
    {
        type     => 'cell',
        format   => $format2,
        criteria => '<',
        value    => 50,
    }
);


###############################################################################
#
# Example 2.
#
$caption = 'Values between 30 and 70 are in light red. '
  . 'Values outside that range are in light green.';

$worksheet2->write( 'A1', $caption );
$worksheet2->write_col( 'B3', $data );

$worksheet2->conditional_formatting( 'B3:K12',
    {
        type     => 'cell',
        format   => $format1,
        criteria => 'between',
        minimum  => 30,
        maximum  => 70,
    }
);

$worksheet2->conditional_formatting( 'B3:K12',
    {
        type     => 'cell',
        format   => $format2,
        criteria => 'not between',
        minimum  => 30,
        maximum  => 70,
    }
);


###############################################################################
#
# Example 3.
#
$caption = 'Duplicate values are in light red. '
  . 'Unique values are in light green.';

$worksheet3->write( 'A1', $caption );
$worksheet3->write_col( 'B3', $data );

# Change a few values to make them unique in the data set.
$worksheet3->write( 'C4', 41 );
$worksheet3->write( 'D8', 51 );
$worksheet3->write( 'I7', 61 );

$worksheet3->conditional_formatting( 'B3:K12',
    {
        type     => 'duplicate',
        format   => $format1,
    }
);

$worksheet3->conditional_formatting( 'B3:K12',
    {
        type     => 'unique',
        format   => $format2,
    }
);


__END__



