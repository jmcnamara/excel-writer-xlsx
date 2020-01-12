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

use Test::More tests => 4;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;
my $format = Excel::Writer::XLSX::Format->new( \{}, \{} );


###############################################################################
#
# Test the _write_conditional_formatting() method.
#
$caption  = " \tWorksheet: _write_conditional_formatting()";
$expected = '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>';

$worksheet = _new_worksheet(\$got);

$worksheet->conditional_formatting( 'A1',
    {
        type     => 'cell',
        format   => $format,
        criteria => 'greater than',
        value    => 5
    }
);

$worksheet->_write_conditional_formats();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_conditional_formatting() method.
#
$caption  = " \tWorksheet: _write_conditional_formatting()";
$expected = '<conditionalFormatting sqref="A2"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>30</formula></cfRule></conditionalFormatting>';

$worksheet = _new_worksheet(\$got);

$worksheet->conditional_formatting( 'A2',
    {
        type     => 'cell',
        format   => $format,
        criteria => 'less than',
        value    => 30
    }
);

$worksheet->_write_conditional_formats();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_conditional_formatting() method.
#
$caption  = " \tWorksheet: _write_conditional_formatting()";
$expected = '<conditionalFormatting sqref="A3"><cfRule type="cellIs" priority="1" operator="greaterThanOrEqual"><formula>50</formula></cfRule></conditionalFormatting>';

$worksheet = _new_worksheet(\$got);

$worksheet->conditional_formatting( 'A3',
    {
        type     => 'cell',
        format   => undef,
        criteria => '>=',
        value    => 50
    }
);

$worksheet->_write_conditional_formats();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_conditional_formatting() method.
#
$caption  = " \tWorksheet: _write_conditional_formatting()";
$expected = '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="between"><formula>10</formula><formula>20</formula></cfRule></conditionalFormatting>';

$worksheet = _new_worksheet(\$got);

$worksheet->conditional_formatting( 'A1',
    {
        type     => 'cell',
        format   => $format,
        criteria => 'between',
        minimum  => 10,
        maximum  => 20,
    }
);

$worksheet->_write_conditional_formats();

is( $got, $expected, $caption );

__END__


