#!/usr/bin/perl -w

##############################################################################
#
# A simple formatting example using Excel::Writer::XLSX.
#
# This program demonstrates the indentation cell format.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#


use strict;
use Excel::Writer::XLSX;

my $workbook = Excel::Writer::XLSX->new( 'indent.xlsx' );

my $worksheet = $workbook->add_worksheet();
my $indent1   = $workbook->add_format( indent => 1 );
my $indent2   = $workbook->add_format( indent => 2 );

$worksheet->set_column( 'A:A', 40 );


$worksheet->write( 'A1', "This text is indented 1 level",  $indent1 );
$worksheet->write( 'A2', "This text is indented 2 levels", $indent2 );

$workbook->close();

__END__
