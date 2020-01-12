###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet data validation methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_got_to_aref _is_deep_diff _new_worksheet);
use strict;
use warnings;

use Test::More tests => 8;


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
# Data validation example 1 from docs.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>0</formula1></dataValidation></dataValidations>';

$worksheet = _new_worksheet(\$got);

$worksheet->data_validation(
    'A1',
    {
        validate => 'integer',
        criteria => '>',
        value    => 0,
    }
);

$worksheet->_write_data_validations();

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Data validation example 1 from docs (with options turned off).
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>0</formula1></dataValidation></dataValidations>';

$worksheet = _new_worksheet(\$got);

$worksheet->data_validation(
    'A1',
    {
        validate     => 'integer',
        criteria     => '>',
        value        => 0,
        ignore_blank => 0,
        show_input   => 0,
        show_error   => 0,
    }
);

$worksheet->_write_data_validations();

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Data validation example 2 from docs.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A2"><formula1>E3</formula1></dataValidation></dataValidations>';

$worksheet = _new_worksheet(\$got);

$worksheet->data_validation(
    'A2',
    {
        validate => 'integer',
        criteria => '>',
        value    => '=E3',
    }
);

$worksheet->_write_data_validations();

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Data validation example 3 from docs.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="decimal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A3"><formula1>0.1</formula1><formula2>0.5</formula2></dataValidation></dataValidations>';

$worksheet = _new_worksheet(\$got);

$worksheet->data_validation(
    'A3',
    {
        validate => 'decimal',
        criteria => 'between',
        minimum  => 0.1,
        maximum  => 0.5,
    }
);

$worksheet->_write_data_validations();

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Data validation example 4 from docs.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A4"><formula1>"open,high,close"</formula1></dataValidation></dataValidations>';

$worksheet = _new_worksheet(\$got);

$worksheet->data_validation(
    'A4',
    {
        validate => 'list',
        source   => [ 'open', 'high', 'close' ],
    }
);

$worksheet->_write_data_validations();

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Data validation example 5 from docs.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A5"><formula1>$E$4:$G$4</formula1></dataValidation></dataValidations>';

$worksheet = _new_worksheet(\$got);

$worksheet->data_validation(
    'A5',
    {
        validate => 'list',
        source   => '=$E$4:$G$4',
    }
);

$worksheet->_write_data_validations();

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Data validation example 6 from docs.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A6"><formula1>39448</formula1><formula2>39794</formula2></dataValidation></dataValidations>';

$worksheet = _new_worksheet(\$got);
$worksheet->{_date_1904} = 0;

$worksheet->data_validation(
    'A6',
    {
        validate => 'date',
        criteria => 'between',
        minimum  => '2008-01-01T',
        maximum  => '2008-12-12T',
    }
);

$worksheet->_write_data_validations();

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Data validation example 7 from docs.
#
$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Enter an integer:" prompt="between 1 and 100" sqref="A7"><formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations>';

$worksheet = _new_worksheet(\$got);

$worksheet->data_validation(
    'A7',
    {
        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 100,
        input_title   => 'Enter an integer:',
        input_message => 'between 1 and 100',
    }
);

$worksheet->_write_data_validations();

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );





__END__
