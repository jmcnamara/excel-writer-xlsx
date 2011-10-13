###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# reverse('(c)'), October 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More;


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
# Test the _write_conditional_formatting() method.
#
$caption  = " \tWorksheet: _write_conditional_formatting()";
$expected = '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>';

$worksheet = _new_worksheet(\$got);

$worksheet->conditional_formatting( 'A1',
    {
        type     => 'cell',
        format   => 1,
        operator => 'greater than',
        formula  => 5
    }
);

$worksheet->_write_conditional_formats();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_conditional_formatting() method.
#
$caption  = " \tWorksheet: _write_conditional_formatting()";
$expected = '<conditionalFormatting sqref="A2"><cfRule type="cellIs" dxfId="1" priority="1" operator="lessThan"><formula>30</formula></cfRule></conditionalFormatting>';

$worksheet = _new_worksheet(\$got);

$worksheet->conditional_formatting( 'A2',
    {
        type     => 'cell',
        format   => 1,
        operator => 'less than',
        formula  => 30
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
        operator => '>=',
        formula  => 50
    }
);

$worksheet->_write_conditional_formats();

is( $got, $expected, $caption );





done_testing();

__END__


