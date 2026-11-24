###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# Copyright 2000-2026, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_workbook);
use strict;
use warnings;

use Test::More tests => 3;



###############################################################################
#
# Test the Workbook _check_sheetname() method.
#
my $tmp;
my $caption = " \tWorkbook: _check_sheetname()";

my $workbook = _new_workbook(\$tmp);
$workbook->add_worksheet();

# All the following add_worksheet() calls should be ignored.

# Test duplicate names.
eval { $workbook->add_worksheet('Sheet1'); };
eval { $workbook->add_worksheet('sheet1'); };

# Test invalid characters.
eval { $workbook->add_worksheet('Sheet['); };
eval { $workbook->add_worksheet('Sheet]'); };
eval { $workbook->add_worksheet('Sheet:'); };
eval { $workbook->add_worksheet('Sheet*'); };
eval { $workbook->add_worksheet('Sheet/'); };
eval { $workbook->add_worksheet('Sheet\\'); };

# Test overly long name.
eval { $workbook->add_worksheet('name_that_is_longer_than_thirty_one_characters'); };

# Test invalid start/end character.
eval { $workbook->add_worksheet(q(Sheet')); };
eval { $workbook->add_worksheet(q('Sheet)); };
eval { $workbook->add_worksheet(q('Sheet')); };


# Test that only 1 worksheet was written.
my $expected = 1;
my $got      = scalar $workbook->sheets();
is($got, $expected, $caption);


###############################################################################
#
# Test the sheetname length limit, which Excel measures in UTF-16 units, so
# characters above U+FFFF, such as emojis, count as 2.
#
$caption  = " \tWorkbook: _check_sheetname() UTF-16 length";
$workbook = _new_workbook(\$tmp);

my $name = "\x{1F600}" x 15 . 'a';
my $worksheet = eval { $workbook->add_worksheet( $name ) };
is( $worksheet && $worksheet->get_name(), $name, $caption );

eval { $workbook->add_worksheet( "\x{1F600}" x 16 ) };
like( $@, qr/must be <= 31 chars/, $caption );



