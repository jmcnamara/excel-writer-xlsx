###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# reverse('©'), September 2012, John McNamara, jmcnamara@cpan.org
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
my $filename     = 'table01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . $filename;
my $exp_filename = $dir . 'xlsx_files/' . $filename;

# TODO
my $ignore_members  = [ 'xl/worksheets/_rels/sheet1.xml.rels' ];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with tables.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();


# TODO

$worksheet->{_tables} = [
    {
        _id         => 1,
        _autofilter => 'C3:F13',
        _columns    => [
            [ 1, 'Column1' ],
            [ 2, 'Column2' ],
            [ 3, 'Column3' ],
            [ 4, 'Column4' ],
        ],
        _style            => 'TableStyleMedium9',
        _show_first_col   => 0,
        _show_last_col    => 0,
        _show_row_stripes => 1,
        _show_col_stripes => 0,
    }
];

$worksheet->set_column('C:F', 10.288);


$worksheet->write('C3', 'Column1');
$worksheet->write('D3', 'Column2');
$worksheet->write('E3', 'Column3');
$worksheet->write('F3', 'Column4');


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



