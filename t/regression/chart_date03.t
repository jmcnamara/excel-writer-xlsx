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
my $filename     = 'chart_date03.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];

# Ignore the default format code for now.
my $ignore_elements = { 'xl/charts/chart1.xml' => [ '<c:formatCode' ] };


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook    = Excel::Writer::XLSX->new( $got_filename );
my $worksheet   = $workbook->add_worksheet();
my $chart       = $workbook->add_chart( type => 'line', embedded => 1 );
my $date_format = $workbook->add_format( num_format => 14 );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids} = [ 51761152, 51762688 ];

$worksheet->set_column('A:A', 12);

my @dates = (
    '2013-01-01T', '2013-01-02T', '2013-01-03T', '2013-01-04T',
    '2013-01-05T', '2013-01-06T', '2013-01-07T', '2013-01-08T',
    '2013-01-09T', '2013-01-10T'
);

my @data = ( 10, 30, 20, 40, 20, 60, 50, 40, 30, 30 );

for my $row ( 0 .. @dates -1 ) {
    $worksheet->write_date_time( $row, 0, $dates[$row], $date_format );
    $worksheet->write( $row, 1, $data[$row] );
}

$chart->add_series(
    categories      => '=Sheet1!$A$1:$A$10',
    values          => '=Sheet1!$B$1:$B$10',
);

$chart->set_x_axis(
    date_axis         => 1,
    minor_unit        => 1,
    major_unit        => 1,
    num_format        => 'dd/mm/yyyy',
    num_format_linked => 1,
);


$worksheet->insert_chart( 'E9', $chart );

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



