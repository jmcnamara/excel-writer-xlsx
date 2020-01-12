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
my $filename     = 'filehandle01.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file as a filehandle.
#
use Excel::Writer::XLSX;

open my $fh, '>', \my $str or die "Failed to open str filehandle: $!";

my $workbook  = Excel::Writer::XLSX->new( $fh );
my $worksheet = $workbook->add_worksheet();

$worksheet->write( 'A1', 'Hello' );
$worksheet->write( 'A2', 123 );

$workbook->close();

open my $out_fh, '>', $got_filename or die "Failed to open out filehandle: $!";
binmode $out_fh;
print $out_fh $str;
close $out_fh;


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



