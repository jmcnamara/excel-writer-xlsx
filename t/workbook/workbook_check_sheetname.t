###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_workbook);
use strict;
use warnings;

use Test::More tests => 1;



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



