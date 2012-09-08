###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# reverse('�'), January 2011, John McNamara, jmcnamara@cpan.org
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
my $filename     = 'chart_line02.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . $filename;
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [];

my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $chart     = $workbook->add_chart( type => 'line', embedded => 1 );

# For testing, copy the randomly generated axis ids in the target xlsx file.
$chart->{_axis_ids}  = [ 63593856, 63612032 ];
$chart->{_axis2_ids} = [ 63615360, 63613568 ];

my $data = [    #
    [ 1, 2, 3, 4, 5 ],
    [ 6, 8, 6, 4, 2 ]
];

$worksheet->write( 'A1', $data );

$chart->add_series( values => '=Sheet1!$A$1:$A$5' );
$chart->add_series(
    values  => '=Sheet1!$B$1:$B$5',
    y2_axis => 1,
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



