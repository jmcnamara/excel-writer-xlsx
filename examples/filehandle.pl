#!/usr/bin/perl

###############################################################################
#
# Example of using Excel::Writer::XLSX to write Excel files to different
# filehandles.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;
use IO::Scalar;


###############################################################################
#
# Example 1. This demonstrates the standard way of creating an Excel file by
# specifying a file name.
#

my $workbook1  = Excel::Writer::XLSX->new( 'fh_01.xlsx' );
my $worksheet1 = $workbook1->add_worksheet();

$worksheet1->write( 0, 0, 'Hi Excel 1' );

$workbook1->close();


###############################################################################
#
# Example 2. Write an Excel file to an existing filehandle.
#

open TEST, '>', 'fh_02.xlsx' or die "Couldn't open file: $!";
binmode TEST;   # Always do this regardless of whether the platform requires it.

my $workbook2  = Excel::Writer::XLSX->new( \*TEST );
my $worksheet2 = $workbook2->add_worksheet();

$worksheet2->write( 0, 0, 'Hi Excel 2' );

$workbook2->close();

###############################################################################
#
# Example 3. Write an Excel file to an existing OO style filehandle.
#

my $fh = FileHandle->new( '> fh_03.xlsx' ) or die "Couldn't open file: $!";

binmode( $fh );

my $workbook3  = Excel::Writer::XLSX->new( $fh );
my $worksheet3 = $workbook3->add_worksheet();

$worksheet3->write( 0, 0, 'Hi Excel 3' );

$workbook3->close();


###############################################################################
#
# Example 4. Write an Excel file to a string via IO::Scalar. Please refer to
# the IO::Scalar documentation for further details.
#

my $xlsx_str;

tie *XLSX, 'IO::Scalar', \$xlsx_str;

my $workbook4  = Excel::Writer::XLSX->new( \*XLSX );
my $worksheet4 = $workbook4->add_worksheet();

$worksheet4->write( 0, 0, 'Hi Excel 4' );
$workbook4->close();    # This is required before we use the scalar


# The Excel file is now in $xlsx_str. As a demonstration, print it to a file.
open TMP, '>', 'fh_04.xlsx' or die "Couldn't open file: $!";
binmode TMP;
print TMP $xlsx_str;
close TMP;


###############################################################################
#
# Example 5. Write an Excel file to a string via IO::Scalar's newer interface.
# Please refer to the IO::Scalar documentation for further details.
#
my $xlsx_str2;

my $fh5 = IO::Scalar->new( \$xlsx_str2 );

my $workbook5  = Excel::Writer::XLSX->new( $fh5 );
my $worksheet5 = $workbook5->add_worksheet();

$worksheet5->write( 0, 0, 'Hi Excel 5' );
$workbook5->close();    # This is required before we use the scalar

# The Excel file is now in $xlsx_str. As a demonstration, print it to a file.
open TMP, '>', 'fh_05.xlsx' or die "Couldn't open file: $!";
binmode TMP;
print TMP $xlsx_str2;
close TMP;

__END__
