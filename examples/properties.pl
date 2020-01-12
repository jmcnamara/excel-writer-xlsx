#!/usr/bin/perl

##############################################################################
#
# An example of adding document properties to a Excel::Writer::XLSX file.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'properties.xlsx' );
my $worksheet = $workbook->add_worksheet();


$workbook->set_properties(
    title    => 'This is an example spreadsheet',
    subject  => 'With document properties',
    author   => 'John McNamara',
    manager  => 'Dr. Heinz Doofenshmirtz',
    company  => 'of Wolves',
    category => 'Example spreadsheets',
    keywords => 'Sample, Example, Properties',
    comments => 'Created with Perl and Excel::Writer::XLSX',
    status   => 'Quo',
);


$worksheet->set_column( 'A:A', 70 );
$worksheet->write( 'A1', qq{Select 'Office Button -> Prepare -> Properties' to see the file properties.} );

$workbook->close();

__END__
