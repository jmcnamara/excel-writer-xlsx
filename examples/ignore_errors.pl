#!/usr/bin/perl -w

#######################################################################
#
# An example of turning off worksheet cells errors/warnings using the
# Excel::Writer::XLSX module.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use strict;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'ignore_errors.xlsx' );
my $worksheet = $workbook->add_worksheet();

# Write strings that looks like numbers. This will cause an Excel warning.
$worksheet->write_string('C2', '123');
$worksheet->write_string('C3', '123');

# Write a divide by zero formula. This will also cause an Excel warning.
$worksheet->write_formula('C5', '=1/0');
$worksheet->write_formula('C6', '=1/0');

# Turn off some of the warnings:
$worksheet->ignore_errors({number_stored_as_text => 'C3', eval_error => 'C6'});

# Write some descriptions for the cells and make the column wider for clarity.
$worksheet->set_column('B:B', 16);
$worksheet->write('B2', 'Warning:');
$worksheet->write('B3', 'Warning turned off:');
$worksheet->write('B5', 'Warning:');
$worksheet->write('B6', 'Warning turned off:');

$workbook->close();

__END__
