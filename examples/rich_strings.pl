#!/usr/bin/perl

#######################################################################
#
# An Excel::Writer::XLSX example showing how to use "rich strings", i.e.,
# strings with multiple formatting.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'rich_strings.xlsx' );
my $worksheet = $workbook->add_worksheet();

$worksheet->set_column( 'A:A', 30 );

# Set some formats to use.
my $bold   = $workbook->add_format( bold        => 1 );
my $italic = $workbook->add_format( italic      => 1 );
my $red    = $workbook->add_format( color       => 'red' );
my $blue   = $workbook->add_format( color       => 'blue' );
my $center = $workbook->add_format( align       => 'center' );
my $super  = $workbook->add_format( font_script => 1 );


# Write some strings with multiple formats.
$worksheet->write_rich_string( 'A1',
    'This is ', $bold, 'bold', ' and this is ', $italic, 'italic' );

$worksheet->write_rich_string( 'A3',
    'This is ', $red, 'red', ' and this is ', $blue, 'blue' );

$worksheet->write_rich_string( 'A5',
    'Some ', $bold, 'bold text', ' centered', $center );

$worksheet->write_rich_string( 'A7',
    $italic, 'j = k', $super, '(n-1)', $center );

$workbook->close();

__END__
