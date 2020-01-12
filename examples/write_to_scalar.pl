#!/usr/bin/perl

##############################################################################
#
# An example of writing an Excel::Writer::XLSX file to a perl scalar.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

# Use a scalar as a filehandle.
open my $fh, '>', \my $str or die "Failed to open filehandle: $!";


# Spreadsheet::WriteExce accepts filehandle as well as file names.
my $workbook  = Excel::Writer::XLSX->new( $fh );
my $worksheet = $workbook->add_worksheet();

$worksheet->write( 0, 0, 'Hi Excel!' );

$workbook->close();


# The Excel file in now in $str. Remember to binmode() the output
# filehandle before printing it.
open my $out_fh, '>', 'write_to_scalar.xlsx'
  or die "Failed to open out filehandle: $!";

binmode $out_fh;
print   $out_fh $str;
close   $out_fh;

__END__

