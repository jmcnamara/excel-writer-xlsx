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
my $priority;
my $param;


###############################################################################
#
# Test the _write_conditional_formatting() method.
#
$caption  = " \tWorksheet: _write_conditional_formatting()";
$expected = '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>';

$priority = 2;
$param = {
    range     => 'A1',
    type      => 'cellIs',
    dxf_index => 1,
    operator  => 'greaterThan',
    formula   => 5
};

$worksheet = _new_worksheet(\$got);

$worksheet->_write_conditional_formatting( $priority, $param );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_conditional_formatting() method.
#
$caption  = " \tWorksheet: _write_conditional_formatting()";
$expected = '<conditionalFormatting sqref="A2"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>30</formula></cfRule></conditionalFormatting>';

$priority = 1;
$param = {
    range     => 'A2',
    type      => 'cellIs',
    dxf_index => 0,
    operator  => 'lessThan',
    formula   => 30
};

$worksheet = _new_worksheet(\$got);

$worksheet->_write_conditional_formatting( $priority, $param );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_conditional_formatting() method.
#
$caption  = " \tWorksheet: _write_conditional_formatting()";
$expected = '<conditionalFormatting sqref="A3"><cfRule type="cellIs" priority="1" operator="greaterThanOrEqual"><formula>50</formula></cfRule></conditionalFormatting>';

$priority = 1;
$param = {
    range     => 'A3',
    type      => 'cellIs',
    dxf_index => undef,
    operator  => 'greaterThanOrEqual',
    formula   => 50
};

$worksheet = _new_worksheet(\$got);

$worksheet->_write_conditional_formatting( $priority, $param );

is( $got, $expected, $caption );





done_testing();

__END__


