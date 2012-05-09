#!/usr/bin/perl -w

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# add shapes and one-to-many connectors to an Excel xlsx file.
#
# reverse('©'), May 2012, John McNamara, jmcnamara@cpan.org
#

use strict;
use Excel::Writer::XLSX;

# Create a new workbook called simple.xls and add a worksheet
my $workbook = Excel::Writer::XLSX->new( 'shape7.xlsx' );

die "Couldn't create new Excel file: $!.\n" unless defined $workbook;

my $worksheet = $workbook->add_worksheet();

# Add a circle, with centered text. c is for circle, not center
my $cw = 60;
my $ch = 60;
my $cx = 210;
my $cy = 190;

my $ellipse = $workbook->add_shape( type => 'ellipse', id=>2, text=>"Hello\nWorld", width=> $cw, height => $ch );
$worksheet->insert_shape('A1', $ellipse, $cx, $cy);

# Add a plus sign at 4 different positions around the circle
my $pw = 20;
my $ph = 20;
my $px = 120;
my $py = 250;
my $plus = $workbook->add_shape( type => 'plus', id=>3, width=> $pw, height => $ph );
my $p1 = $worksheet->insert_shape('A1', $plus, 350, 350);
my $p2 = $worksheet->insert_shape('A1', $plus, 150, 350);
my $p3 = $worksheet->insert_shape('A1', $plus, 350, 150);
$plus->{adjustments} = [35];    # change shape of plus symbol
my $p4 = $worksheet->insert_shape('A1', $plus, 150, 150);

my $cxn_shape = $workbook->add_shape( type => 'bentConnector3', fill=> 0);

$cxn_shape->{start} = $ellipse->{id};
$cxn_shape->{start_idx} = 4;                # 4th connection from top
$cxn_shape->{start_side} = 'b';             # b)ottom

$cxn_shape->{end} = $p1->{id};
$cxn_shape->{end_idx} = 0;                  # first connection (zero based) from top
$cxn_shape->{end_side} = 't';               # t)op
$worksheet->insert_shape('A1', $cxn_shape, 0, 0);

$cxn_shape->{end} = $p2->{id};
$worksheet->insert_shape('A1', $cxn_shape, 0, 0);

$cxn_shape->{end} = $p3->{id};
$worksheet->insert_shape('A1', $cxn_shape, 0, 0);

$cxn_shape->{end} = $p4->{id};
$cxn_shape->{adjustments} = [-50, 45, 120];    
$worksheet->insert_shape('A1', $cxn_shape, 0, 0);


