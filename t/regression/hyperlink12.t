###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
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
my $filename     = 'hyperlink12.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members = [];

my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with hyperlinks.
# This example has link formatting.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();
my $format    = $workbook->add_format( color => 'blue', underline => 1 );

$worksheet->write_url( 'A1', 'mailto:jmcnamara@cpan.org', $format );

$worksheet->write_url( 'A3', 'ftp://perl.org/', $format );

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



