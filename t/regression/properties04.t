###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# reverse ('(c)'), April 2011, John McNamara, jmcnamara@cpan.org
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
my $filename     = 'properties04.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = { 'xl/workbook.xml' => ['<workbookView'] };


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();


my $long_string = 'This is a long string. ' x 11;
$long_string .= 'AA';

$workbook->set_custom_property( 'Checked by',      'Adam'                             );
$workbook->set_custom_property( 'Date completed',  '2016-12-12T23:00:00Z', 'date'     );
$workbook->set_custom_property( 'Document number', '12345' ,               'num_int'  );
$workbook->set_custom_property( 'Reference',       '1.2345',               'num_real' );
$workbook->set_custom_property( 'Source',          1,                      'bool'     );
$workbook->set_custom_property( 'Status',          0,                      'bool'     );
$workbook->set_custom_property( 'Department',      $long_string                       );
$workbook->set_custom_property( 'Group',           '1.2345678901234',      'num_real' );

$worksheet->set_column( 'A:A', 70 );
$worksheet->write( 'A1', qq{Select 'Office Button -> Prepare -> Properties' to see the file properties.} );

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



