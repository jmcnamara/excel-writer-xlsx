#!/usr/bin/perl

##############################################################################
#
# A simple example of writing some Russian cyrillic text using
# Excel::Writer::XLSX.
#
#
#
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;


# In this example we generate utf8 strings from character data but in a
# real application we would expect them to come from an external source.
#


# Create a Russian worksheet name in utf8.
my $sheet = pack "U*", 0x0421, 0x0442, 0x0440, 0x0430, 0x043D, 0x0438,
  0x0446, 0x0430;


# Create a Russian string.
my $str = pack "U*", 0x0417, 0x0434, 0x0440, 0x0430, 0x0432, 0x0441,
  0x0442, 0x0432, 0x0443, 0x0439, 0x0020, 0x041C,
  0x0438, 0x0440, 0x0021;


my $workbook = Excel::Writer::XLSX->new( 'unicode_cyrillic.xlsx' );

die "Couldn't create new Excel file: $!.\n" unless defined $workbook;

my $worksheet = $workbook->add_worksheet( $sheet . '1' );

$worksheet->set_column( 'A:A', 18 );
$worksheet->write( 'A1', $str );

$workbook->close();

__END__

