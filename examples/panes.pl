#!/usr/bin/perl

#######################################################################
#
# Example of using the Excel::Writer::XLSX module to create worksheet panes.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook = Excel::Writer::XLSX->new( 'panes.xlsx' );

my $worksheet1 = $workbook->add_worksheet( 'Panes 1' );
my $worksheet2 = $workbook->add_worksheet( 'Panes 2' );
my $worksheet3 = $workbook->add_worksheet( 'Panes 3' );
my $worksheet4 = $workbook->add_worksheet( 'Panes 4' );

# Freeze panes
$worksheet1->freeze_panes( 1, 0 );    # 1 row

$worksheet2->freeze_panes( 0, 1 );    # 1 column
$worksheet3->freeze_panes( 1, 1 );    # 1 row and column

# Split panes.
# The divisions must be specified in terms of row and column dimensions.
# The default row height is 15 and the default column width is 8.43
#
$worksheet4->split_panes( 15, 8.43 );    # 1 row and column


#######################################################################
#
# Set up some formatting and text to highlight the panes
#

my $header = $workbook->add_format(
    align    => 'center',
    valign   => 'vcenter',
    fg_color => '#C3FFC0',
);

my $center = $workbook->add_format( align => 'center' );


#######################################################################
#
# Sheet 1
#

$worksheet1->set_column( 'A:I', 16 );
$worksheet1->set_row( 0, 20 );
$worksheet1->set_selection( 'C3' );

for my $i ( 0 .. 8 ) {
    $worksheet1->write( 0, $i, 'Scroll down', $header );
}

for my $i ( 1 .. 100 ) {
    for my $j ( 0 .. 8 ) {
        $worksheet1->write( $i, $j, $i + 1, $center );
    }
}


#######################################################################
#
# Sheet 2
#

$worksheet2->set_column( 'A:A', 16 );
$worksheet2->set_selection( 'C3' );

for my $i ( 0 .. 49 ) {
    $worksheet2->set_row( $i, 15 );
    $worksheet2->write( $i, 0, 'Scroll right', $header );
}

for my $i ( 0 .. 49 ) {
    for my $j ( 1 .. 25 ) {
        $worksheet2->write( $i, $j, $j, $center );
    }
}


#######################################################################
#
# Sheet 3
#

$worksheet3->set_column( 'A:Z', 16 );
$worksheet3->set_selection( 'C3' );

$worksheet3->write( 0, 0, '', $header );

for my $i ( 1 .. 25 ) {
    $worksheet3->write( 0, $i, 'Scroll down', $header );
}

for my $i ( 1 .. 49 ) {
    $worksheet3->write( $i, 0, 'Scroll right', $header );
}

for my $i ( 1 .. 49 ) {
    for my $j ( 1 .. 25 ) {
        $worksheet3->write( $i, $j, $j, $center );
    }
}


#######################################################################
#
# Sheet 4
#

$worksheet4->set_selection( 'C3' );

for my $i ( 1 .. 25 ) {
    $worksheet4->write( 0, $i, 'Scroll', $center );
}

for my $i ( 1 .. 49 ) {
    $worksheet4->write( $i, 0, 'Scroll', $center );
}

for my $i ( 1 .. 49 ) {
    for my $j ( 1 .. 25 ) {
        $worksheet4->write( $i, $j, $j, $center );
    }
}

$workbook->close();

__END__
