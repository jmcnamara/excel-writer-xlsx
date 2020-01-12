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
my $filename     = 'quote_name03.xlsx';
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

my $data = [
    [ 1, 2, 3, 4,  5 ],
    [ 2, 4, 6, 8,  10 ],
    [ 3, 6, 9, 12, 15 ],

];


# Test quoted/non-quoted sheet names.
my @sheetnames = (
    'Sheet<1', 'Sheet>2', 'Sheet=3', 'Sheet@4',
    'Sheet^5', 'Sheet`6', 'Sheet_7', 'Sheet~8'
);

for my $sheetname ( @sheetnames ) {

    my $worksheet = $workbook->add_worksheet( $sheetname );
    my $chart = $workbook->add_chart( type => 'pie', embedded => 1 );

    $worksheet->write( 'A1', $data );
    $chart->add_series(values => [$sheetname, 0, 4, 0, 0]);
    $worksheet->insert_chart( 'E6', $chart, 26, 17 );

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

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Cleanup.
#
unlink $got_filename;

__END__
