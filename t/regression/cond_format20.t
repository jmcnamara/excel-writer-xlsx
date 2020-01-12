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
my $filename     = 'cond_format20.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = { 'xl/workbook.xml' => ['<workbookView'], };


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with conditional
# formatting.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

$worksheet->write( 'A1', 10 );
$worksheet->write( 'A2', 20 );
$worksheet->write( 'A3', 30 );
$worksheet->write( 'A4', 40 );

$worksheet->conditional_formatting( 'A1:A4',
    {
        type         => 'icon_set',
        icon_style   => '3_arrows',
        icons        => [ {criteria => '>',  type => 'percent', value => 0},
                          {criteria => '<',  type => 'percent', value => 0},
                          {criteria => '>=', type => 'percent', value => 0},
                        ],
    }
);


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
