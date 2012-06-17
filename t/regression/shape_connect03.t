###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# reverse('©'), January 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_compare_xlsx_files _is_deep_diff);
use strict;
use warnings;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $filename     = 'shape_connect03.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . $filename;
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};

###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
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

$workbook->close();

###############################################################################
#
# Compare the generated and existing Excel files.
#

my ( $got, $expected, $caption ) = _compare_xlsx_files(

    $got_filename,
    $exp_filename,
    $ignore_members,
    $ignore_elements,
);

$caption .= ' # connected shapes 4 connections';
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Cleanup.
#
unlink $got_filename;

__END__



