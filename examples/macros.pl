#!/usr/bin/perl

#######################################################################
#
# An example of adding macros to an Excel::Writer::XLSX file using
# a VBA project file extracted from an existing Excel xlsm file.
#
# The C<extract_vba> utility supplied with Excel::Writer::XLSX can be
# used to extract the vbaProject.bin file.
#
# An embedded macro is connected to a form button on the worksheet.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

# Note the file extension should be .xlsm.
my $workbook  = Excel::Writer::XLSX->new( 'macros.xlsm' );
my $worksheet = $workbook->add_worksheet();

$worksheet->set_column( 'A:A', 30 );

# Add the VBA project binary.
$workbook->add_vba_project( './vbaProject.bin' );

# Show text for the end user.
$worksheet->write( 'A3', 'Press the button to say hello.' );

# Add a button tied to a macro in the VBA project.
$worksheet->insert_button(
    'B3',
    {
        macro   => 'say_hello',
        caption => 'Press Me',
        width   => 80,
        height  => 30
    }
);

$workbook->close();

__END__
