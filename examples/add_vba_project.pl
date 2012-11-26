#!/usr/bin/perl

#######################################################################
#
# An example of adding macros to an Excel::Writer::XLSX file using
# a VBA project file extracted from an existing Excel xlsm file.
#
# The C<extract_vba> utility supplied with Excel::Writer::XLSX can be
# used to extract the vbaProject.bin file.
#
# reverse('(c)'), November 2012, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

# Note the file extension should be .xlsm.
my $workbook  = Excel::Writer::XLSX->new( 'add_vba_project.xlsm' );
my $worksheet = $workbook->add_worksheet();

$worksheet->set_column( 'A:A', 50 );

# Add the VBA project binary.
$workbook->add_vba_project( './vbaProject.bin' );

# Show text for the end user.
$worksheet->write( 'A1', 'Run the SampleMacro embedded in this file.' );
$worksheet->write( 'A2', 'You may have to turn on the Excel Developer option first.' );

# Call a user defined function from the VBA project.
$worksheet->write( 'A6', 'Result from a user defined function:' );
$worksheet->write( 'B6', '=MyFunction(7)' );



__END__
