#!/usr/bin/perl

########################################################################
#
# Example of cell locking and formula hiding in an Excel worksheet via
# the Excel::Writer::XLSX module.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'protection.xlsx' );
my $worksheet = $workbook->add_worksheet();

# Create some format objects
my $unlocked = $workbook->add_format( locked => 0 );
my $hidden   = $workbook->add_format( hidden => 1 );

# Format the columns
$worksheet->set_column( 'A:A', 45 );
$worksheet->set_selection( 'B3' );

# Protect the worksheet
$worksheet->protect();

# Examples of cell locking and hiding.
$worksheet->write( 'A1', 'Cell B1 is locked. It cannot be edited.' );
$worksheet->write_formula( 'B1', '=1+2', undef, 3 );    # Locked by default.

$worksheet->write( 'A2', 'Cell B2 is unlocked. It can be edited.' );
$worksheet->write_formula( 'B2', '=1+2', $unlocked, 3 );

$worksheet->write( 'A3', "Cell B3 is hidden. The formula isn't visible." );
$worksheet->write_formula( 'B3', '=1+2', $hidden, 3 );

$worksheet->write( 'A5', 'Use Menu->Tools->Protection->Unprotect Sheet' );
$worksheet->write( 'A6', 'to remove the worksheet protection.' );

$workbook->close();

__END__

