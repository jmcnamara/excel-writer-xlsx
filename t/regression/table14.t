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
my $filename     = 'table14.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;


my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with tables.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $format1   = $workbook->add_format( num_format => '0.00;[Red]0.00',       dxf_index => 2 );
my $format2   = $workbook->add_format( num_format => '0.00_ ;\-0.00\ ',      dxf_index => 1 );
my $format3   = $workbook->add_format( num_format => '0.00_ ;[Red]\-0.00\ ', dxf_index => 0 );

my $data = [
    [ 'Foo', 1234, 2000, 4321 ],
    [ 'Bar', 1256, 4000, 4320 ],
    [ 'Baz', 2234, 3000, 4332 ],
    [ 'Bop', 1324, 1000, 4333 ],
];


# Set the column width to match the target worksheet.
$worksheet->set_column('C:F', 10.288);

# Add the table.
$worksheet->add_table(
    'C2:F6',
    {
        data    => $data,
        columns => [
            {},
            { format => $format1 },
            { format => $format2 },
            { format => $format3 },
        ]
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



