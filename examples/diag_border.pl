#!/usr/bin/perl -w

##############################################################################
#
# A simple formatting example that demonstrates how to add a diagonal cell
# border with Excel::Writer::XLSX
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;


my $workbook  = Excel::Writer::XLSX->new( 'diag_border.xlsx' );
my $worksheet = $workbook->add_worksheet();


my $format1 = $workbook->add_format( diag_type => 1 );

my $format2 = $workbook->add_format( diag_type => 2 );

my $format3 = $workbook->add_format( diag_type => 3 );

my $format4 = $workbook->add_format(
    diag_type   => 3,
    diag_border => 7,
    diag_color  => 'red',
);


$worksheet->write( 'B3',  'Text', $format1 );
$worksheet->write( 'B6',  'Text', $format2 );
$worksheet->write( 'B9',  'Text', $format3 );
$worksheet->write( 'B12', 'Text', $format4 );

$workbook->close();

__END__

