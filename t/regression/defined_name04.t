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
use utf8;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $filename     = 'defined_name04.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [];

my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with defined names.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

# Test for valid Excel defined names.
$workbook->define_name( q(\\__),     q(=Sheet1!$A$1) );
$workbook->define_name( q(a3f6),     q(=Sheet1!$A$2) );
$workbook->define_name( q(afoo.bar), q(=Sheet1!$A$3) );
$workbook->define_name( q(étude),    q(=Sheet1!$A$4) );
$workbook->define_name( q(eésumé),   q(=Sheet1!$A$5) );
$workbook->define_name( q(a),        q(=Sheet1!$A$6) );

# The following are not valid Excel names and shouldn't be written to
# the output file. We also catch the warnings and ignore them.
local $SIG{__WARN__} = sub {};
eval {
    $workbook->define_name( q(.abc),       q(=Sheet1!$B$1) );
    $workbook->define_name( q(GFG$),       q(=Sheet1!$B$1) );
    $workbook->define_name( q(A1),         q(=Sheet1!$B$1) );
    $workbook->define_name( q(XFD1048576), q(=Sheet1!$B$1) );
    $workbook->define_name( q(1A),         q(=Sheet1!$B$1) );
    $workbook->define_name( q(A A),        q(=Sheet1!$B$1) );
    $workbook->define_name( q(c),          q(=Sheet1!$B$1) );
    $workbook->define_name( q(r),          q(=Sheet1!$B$1) );
    $workbook->define_name( q(C),          q(=Sheet1!$B$1) );
    $workbook->define_name( q(R),          q(=Sheet1!$B$1) );
    $workbook->define_name( q(R1),         q(=Sheet1!$B$1) );
    $workbook->define_name( q(C1),         q(=Sheet1!$B$1) );
    $workbook->define_name( q(R1C1),       q(=Sheet1!$B$1) );
    $workbook->define_name( q(R13C99),     q(=Sheet1!$B$1) );
};

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



