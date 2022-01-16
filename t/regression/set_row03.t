###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2022, John McNamara, jmcnamara@cpan.org
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
my $filename     = 'set_row03.xlsx';
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

$worksheet->set_row( 0,  0.75  );
$worksheet->set_row( 1,  1.50  );
$worksheet->set_row( 2,  2.25  );
$worksheet->set_row( 3,  3     );

$worksheet->set_row( 11, 9     );
$worksheet->set_row( 12, 9.75  );
$worksheet->set_row( 13, 10.50 );
$worksheet->set_row( 14, 11.25 );

$worksheet->set_row( 18, 14.25 );
$worksheet->set_row( 20, 15.75, undef, 1 );
$worksheet->set_row( 21, 16.50 );

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



