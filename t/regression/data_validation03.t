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
my $filename     = 'data_validation03.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of an Excel::Writer::XLSX file data validation.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

$worksheet->data_validation(
    'C2',
    {
        validate      => 'list',
        value         => [ 'Foo', 'Bar', 'Baz' ],
        input_title   => 'This is the input title',
        input_message => 'This is the input message',
    }
);

# Examples of the maximum input.
my $input_title   = 'This is the longest input title1';
my $input_message = 'This is the longest input message ' . ('a' x 221);
my $values = [
    "Foobar", "Foobas", "Foobat", "Foobau", "Foobav", "Foobaw", "Foobax",
    "Foobay", "Foobaz", "Foobba", "Foobbb", "Foobbc", "Foobbd", "Foobbe",
    "Foobbf", "Foobbg", "Foobbh", "Foobbi", "Foobbj", "Foobbk", "Foobbl",
    "Foobbm", "Foobbn", "Foobbo", "Foobbp", "Foobbq", "Foobbr", "Foobbs",
    "Foobbt", "Foobbu", "Foobbv", "Foobbw", "Foobbx", "Foobby", "Foobbz",
    "Foobca", "End"
];

$worksheet->data_validation(
    'D6',
    {
        validate      => 'list',
        value         => $values,
        input_title   => $input_title,
        input_message => $input_message,
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



