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
my $filename     = 'row_col_format09.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test Excel::Writer::XLSX file with row or column formatting.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $bold      = $workbook->add_format( bold => 1 );
my $mixed     = $workbook->add_format( bold => 1 , italic => 1 );
my $italic    = $workbook->add_format( italic => 1 );

$workbook->_set_default_xf_indices();

$worksheet->set_row( 4, undef, $bold );
$worksheet->set_column( 'C:C', undef, $italic );

$worksheet->write( 'C1', 'Foo' );
$worksheet->write( 'A5', 'Foo' );
$worksheet->write( 'C5', 'Foo', $mixed );

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

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Cleanup.
#
unlink $got_filename;

__END__



