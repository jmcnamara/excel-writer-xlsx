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
my $worksheet4 = $workbook->add_worksheet();


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
    [ 34, 72,  38, 30, 75, 48, 75, 66, 84, 86 ],
    [ 6,  24,  1,  84, 54, 62, 60, 3,  26, 59 ],
    [ 28, 79,  97, 13, 85, 93, 93, 22, 5,  14 ],
    [ 27, 71,  40, 17, 18, 79, 90, 93, 29, 47 ],
    [ 88, 25,  33, 23, 67, 1,  59, 79, 47, 36 ],
    [ 24, 100, 20, 88, 29, 33, 38, 54, 54, 88 ],
    [ 6,  57,  88, 28, 10, 26, 37, 7,  41, 48 ],
    [ 52, 78,  1,  96, 26, 45, 47, 33, 96, 36 ],
    [ 60, 54,  81, 66, 81, 90, 80, 93, 12, 55 ],
    [ 70, 5,   46, 14, 71, 19, 66, 36, 41, 21 ],
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
        criteria => '>=',
        value    => 50,
        format   => $format1,
    }
);

# Write another conditional format over the same range.
$worksheet1->conditional_formatting( 'B3:K12',
    {
        type     => 'cell',
        criteria => '<',
        value    => 50,
        format   => $format2,
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
        criteria => 'between',
        minimum  => 30,
        maximum  => 70,
        format   => $format1,
    }
);

$worksheet2->conditional_formatting( 'B3:K12',
    {
        type     => 'cell',
        criteria => 'not between',
        minimum  => 30,
        maximum  => 70,
        format   => $format2,
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


###############################################################################
#
# Example 4.
#
$caption = 'Above average values are in light red. '
  . 'Below average values are in light green.';

$worksheet4->write( 'A1', $caption );
$worksheet4->write_col( 'B3', $data );

# Change a few values to make them unique in the data set.
$worksheet4->write( 'C4', 41 );
$worksheet4->write( 'D8', 51 );
$worksheet4->write( 'I7', 61 );

$worksheet4->conditional_formatting( 'B3:K12',
    {
        type     => 'average',
        criteria => 'above',
        format   => $format1,
    }
);

$worksheet4->conditional_formatting( 'B3:K12',
    {
        type     => 'average',
        criteria => 'below',
        format   => $format2,
    }
);


__END__



