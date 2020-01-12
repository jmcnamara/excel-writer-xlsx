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
my $filename     = 'shape_scale01.xlsx';
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
my $normal = $workbook->add_shape( 
    name => 'chip', 
    type => 'diamond', 
    text=> "Normal", 
    width=> 100, 
    height => 100,
    format => $format,
);

$worksheet->insert_shape('A1', $normal,  50, 50);

$normal->set_text('Scaled 3w x 2h');
$normal->set_name('Hope');
$worksheet->insert_shape('A1', $normal, 250, 50, 3, 2 );

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

$caption .= ' # scaled shapes';
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Cleanup.
#
unlink $got_filename;

__END__



