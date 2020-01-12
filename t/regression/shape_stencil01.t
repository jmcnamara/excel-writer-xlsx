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
my $filename     = 'shape_stencil01.xlsx';
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
$worksheet->hide_gridlines(2);

my $format = $workbook->add_format(font => 'Arial', size => 8);
my $shape = $workbook->add_shape( 
    type => 'rect', 
    width=> 90, 
    height => 90,
    format => $format,
);

for my $n (1..10) {
    # Change the last 5 rectangles to stars.  Previously inserted shapes stay as rectangles
    $shape->set_type('star5') if $n == 6;
    my $text = $shape->get_type(); 
    $shape->set_text( join (' ', $text, $n) ); 
    $worksheet->insert_shape('A1', $shape,  $n * 100, 50);
}

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

$caption .= ' # stencils';
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Cleanup.
#
unlink $got_filename;

__END__



