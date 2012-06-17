#!/usr/bin/perl -w

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# add shapes (objects and top/bottom connectors) to an Excel xlsx file.
#
# reverse('©'), May 2012, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

# Create a new workbook called simple.xls and add a worksheet
my $workbook = Excel::Writer::XLSX->new( 'shape5.xlsx' );

die "Couldn't create new Excel file: $!.\n" unless defined $workbook;

my $worksheet = $workbook->add_worksheet();

my $s1 = $workbook->add_shape( type => 'ellipse', width=> 60, height => 60 );
$worksheet->insert_shape('A1', $s1, 50, 50);

my $s2 = $workbook->add_shape( type => 'plus', width=> 20, height => 20 );
$worksheet->insert_shape('A1', $s2, 250, 200);

# Create a connector to link the two shapes
my $cxn_shape = $workbook->add_shape( type => 'bentConnector3' );

# Link the start of the connector to the right side
$cxn_shape->{start} = $s1->{id};
$cxn_shape->{start_idx} = 4;                # 4th connection point, clockwise from top(0)
$cxn_shape->{start_side} = 'b';             # r)ight or b)ottom

# Link the end of the connector to the left side
$cxn_shape->{end} = $s2->{id};
$cxn_shape->{end_idx} = 0;                  # 4th connection point, clockwise from top(0)
$cxn_shape->{end_side} = 't';               # l)eft or t)op

$worksheet->insert_shape('A1', $cxn_shape, 0, 0);

# `shape5.xlsx`;
__END__
