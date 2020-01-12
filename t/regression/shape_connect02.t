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
my $filename     = 'shape_connect02.xlsx';
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

my $s1 = $workbook->add_shape( type => 'chevron', width=> 60, height => 60 );
$worksheet->insert_shape('A1', $s1, 50, 50);

my $s2 = $workbook->add_shape( type => 'pentagon', width=> 20, height => 20 );
$worksheet->insert_shape('A1', $s2, 250, 200);

# Create a connector to link the two shapes
my $cxn_shape = $workbook->add_shape( type => 'curvedConnector3' );

# Link the start of the connector to the right side
$cxn_shape->set_start( $s1->get_id() );
$cxn_shape->set_start_index(2);                # 2nd connection point, clockwise from top(0)
$cxn_shape->set_start_side('r');               # r)ight or b)ottom

# Link the end of the connector to the left side
$cxn_shape->set_end ( $s2->get_id() );
$cxn_shape->set_end_index(4);                  # 4th connection point, clockwise from top(0)
$cxn_shape->set_end_side('l');                 # l)eft or t)op

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

$caption .= ' # connected shapes r/l';
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Cleanup.
#
unlink $got_filename;

__END__



