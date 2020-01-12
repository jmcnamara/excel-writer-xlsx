###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
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
my $filename     = 'shape_connect01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
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
my $format = $workbook->add_format(font => 'Arial', size => 8);

# Add a circle, with centered text
my $ellipse = $workbook->add_shape( type => 'ellipse', text=>"Hello\nWorld", width=> 60, height => 60, format => $format);
$worksheet->insert_shape('A1', $ellipse, 50, 50);

# Add a plus
my $plus = $workbook->add_shape( type => 'plus', width=> 20, height => 20 );
$worksheet->insert_shape('A1', $plus, 250, 200);

# Create a bent connector to link the two shapes
my $cxn_shape = $workbook->add_shape( type => 'bentConnector3' );

# Link the connector to the bottom of the circle
$cxn_shape->set_start( $ellipse->get_id() );
$cxn_shape->set_start_index(4);                # 4th connection point, clockwise from top(0)
$cxn_shape->set_start_side('b');               # r)ight or b)ottom

# Link the connector to the bottom of the plus sign
$cxn_shape->set_end ( $plus->get_id() );
$cxn_shape->set_end_index(0);                  # 0 - top connection point
$cxn_shape->set_end_side('t');                 # l)eft or t)op

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

$caption .= ' # connected shapes t/b';
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Cleanup.
#
unlink $got_filename;

__END__



