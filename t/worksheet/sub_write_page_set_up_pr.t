###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 1;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;

###############################################################################
#
# Test the _write_page_set_up_pr() method.
#
$caption  = " \tWorksheet: _write_page_set_up_pr()";
$expected = '<pageSetUpPr fitToPage="1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->{_fit_page} = 1;

$worksheet->_write_page_set_up_pr();

is( $got, $expected, $caption );

__END__


