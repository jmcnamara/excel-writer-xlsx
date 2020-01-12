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

use Test::More tests => 48;


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
# Test 1 Integer between 1 and 10.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 10,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: integer between";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 2 Integer not between 1 and 10.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate => 'integer',
        criteria => 'not between',
        minimum  => 1,
        maximum  => 10,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: integer not between";
$expected = '<dataValidations count="1"><dataValidation type="whole" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test 3,4,5 Integer == 1.
#
for my $operator ( 'equal to', '=', '==' ) {
    $worksheet = _new_worksheet( \$got );

    $worksheet->data_validation(
        'B5',
        {
            validate => 'integer',
            criteria => $operator,
            value    => 1,
        }
    );

    $worksheet->_write_data_validations();

    $caption  = " \tData validation api: integer equal to";
    $expected = '<dataValidations count="1"><dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>';

    $expected = _got_to_aref( $expected );
    $got      = _got_to_aref( $got );
    _is_deep_diff( $got, $expected, $caption );
}


###############################################################################
#
# Test 6,7,8 Integer != 1.
#
for my $operator ( 'not equal to', '<>', '!=' ) {
    $worksheet = _new_worksheet( \$got );

    $worksheet->data_validation(
        'B5',
        {
            validate => 'integer',
            criteria => $operator,
            value    => 1,
        }
    );

    $worksheet->_write_data_validations();

    $caption  = " \tData validation api: integer not equal to";
    $expected = '<dataValidations count="1"><dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>';

    $expected = _got_to_aref( $expected );
    $got      = _got_to_aref( $got );
    _is_deep_diff( $got, $expected, $caption );
}


###############################################################################
#
# Test 9,10 Integer > 1.
#
for my $operator ( 'greater than', '>' ) {
    $worksheet = _new_worksheet( \$got );

    $worksheet->data_validation(
        'B5',
        {
            validate => 'integer',
            criteria => $operator,
            value    => 1,
        }
    );

    $worksheet->_write_data_validations();

    $caption  = " \tData validation api: integer >";
    $expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>';

    $expected = _got_to_aref( $expected );
    $got      = _got_to_aref( $got );
    _is_deep_diff( $got, $expected, $caption );
}


###############################################################################
#
# Test 11,12 Integer < 1.
#
for my $operator ( 'less than', '<' ) {
    $worksheet = _new_worksheet( \$got );

    $worksheet->data_validation(
        'B5',
        {
            validate => 'integer',
            criteria => $operator,
            value    => 1,
        }
    );

    $worksheet->_write_data_validations();

    $caption  = " \tData validation api: integer <";
    $expected = '<dataValidations count="1"><dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>';

    $expected = _got_to_aref( $expected );
    $got      = _got_to_aref( $got );
    _is_deep_diff( $got, $expected, $caption );
}


###############################################################################
#
# Test 13,14 Integer >= 1.
#
for my $operator ( 'greater than or equal to', '>=' ) {
    $worksheet = _new_worksheet( \$got );

    $worksheet->data_validation(
        'B5',
        {
            validate => 'integer',
            criteria => $operator,
            value    => 1,
        }
    );

    $worksheet->_write_data_validations();

    $caption  = " \tData validation api: integer >=";
    $expected = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>';

    $expected = _got_to_aref( $expected );
    $got      = _got_to_aref( $got );
    _is_deep_diff( $got, $expected, $caption );
}


###############################################################################
#
# Test 15,16 Integer <= 1.
#
for my $operator ( 'less than or equal to', '<=' ) {
    $worksheet = _new_worksheet( \$got );

    $worksheet->data_validation(
        'B5',
        {
            validate => 'integer',
            criteria => $operator,
            value    => 1,
        }
    );

    $worksheet->_write_data_validations();

    $caption  = " \tData validation api: integer <=";
    $expected = '<dataValidations count="1"><dataValidation type="whole" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1></dataValidation></dataValidations>';

    $expected = _got_to_aref( $expected );
    $got      = _got_to_aref( $got );
    _is_deep_diff( $got, $expected, $caption );
}


###############################################################################
#
# Test 17 Integer between 1 and 10 (same as test 1) + Ignore blank off.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate     => 'integer',
        criteria     => 'between',
        minimum      => 1,
        maximum      => 10,
        ignore_blank => 0,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: ignore blank off";
$expected = '<dataValidations count="1"><dataValidation type="whole" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 18 Integer between 1 and 10 (same as test 1) + Error style == warning..
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate   => 'integer',
        criteria   => 'between',
        minimum    => 1,
        maximum    => 10,
        error_type => 'warning',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: error style = warning";
$expected = '<dataValidations count="1"><dataValidation type="whole" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 19 Integer between 1 and 10 (same as test 1) + Error style == infor..
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate   => 'integer',
        criteria   => 'between',
        minimum    => 1,
        maximum    => 10,
        error_type => 'information',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: error style = information";
$expected = '<dataValidations count="1"><dataValidation type="whole" errorStyle="information" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 20 Integer between 1 and 10 (same as test 1)
#         + input title.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate    => 'integer',
        criteria    => 'between',
        minimum     => 1,
        maximum     => 10,
        input_title => 'Input title January',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: with input title";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Input title January" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test 21 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 10,
        input_title   => 'Input title January',
        input_message => 'Input message February',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api:   + input message";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 22 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#         + error title.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 10,
        input_title   => 'Input title January',
        input_message => 'Input message February',
        error_title   => 'Error title March',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api:   + error title";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Error title March" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 23 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#         + error title.
#         + error message.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 10,
        input_title   => 'Input title January',
        input_message => 'Input message February',
        error_title   => 'Error title March',
        error_message => 'Error message April',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api:   + error message";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Error title March" error="Error message April" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 24 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#         + error title.
#         + error message.
#         - input message box.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 10,
        input_title   => 'Input title January',
        input_message => 'Input message February',
        error_title   => 'Error title March',
        error_message => 'Error message April',
        show_input    => 0,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: no input box";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showErrorMessage="1" errorTitle="Error title March" error="Error message April" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 25 Integer between 1 and 10 (same as test 1)
#         + input title.
#         + input message.
#         + error title.
#         + error message.
#         - input message box.
#         - error message box.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate      => 'integer',
        criteria      => 'between',
        minimum       => 1,
        maximum       => 10,
        input_title   => 'Input title January',
        input_message => 'Input message February',
        error_title   => 'Error title March',
        error_message => 'Error message April',
        show_input    => 0,
        show_error    => 0,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: no error box";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" errorTitle="Error title March" error="Error message April" promptTitle="Input title January" prompt="Input message February" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test 26 'Any' value on its own shouldn't produce a DV record.
#
$got = '';
$worksheet = _new_worksheet( \$got );

$worksheet->data_validation( 'B5', { validate => 'any', } );

$worksheet->_write_data_validations();

$caption  = " \tData validation api: any validation";
$expected = '';

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 27 Decimal = 1.2345
#
$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate => 'decimal',
        criteria => '==',
        value    => 1.2345,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: decimal validation";
$expected = '<dataValidations count="1"><dataValidation type="decimal" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1.2345</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test 28 List = a,bb,ccc
#
$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate => 'list',
        source   => [ 'a', 'bb', 'ccc' ],
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: explicit list";
$expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>"a,bb,ccc"</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 29 List = a,bb,ccc, No dropdown
#
$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate => 'list',
        source   => [ 'a', 'bb', 'ccc' ],
        dropdown => 0,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: list with no dropdown";
$expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>"a,bb,ccc"</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 30 List = $D$1:$D$5
#
$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'A1:A1',
    {
        validate => 'list',
        source   => '=$D$1:$D$5',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: list with range";
$expected = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>$D$1:$D$5</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test 31 Date = 39653 (2008-07-24)
#
$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate => 'date',
        criteria => '==',
        value    => 39653,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: date";
$expected = '<dataValidations count="1"><dataValidation type="date" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>39653</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test 32 Date = 2008-07-24T
#
$worksheet = _new_worksheet( \$got );
$worksheet->{_date_1904} = 0;

$worksheet->data_validation(
    'B5',
    {
        validate => 'date',
        criteria => '==',
        value    => '2008-07-24T',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: date auto";
$expected = '<dataValidations count="1"><dataValidation type="date" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>39653</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test 33 Date between ranges.
#
$worksheet = _new_worksheet( \$got );
$worksheet->{_date_1904} = 0;

$worksheet->data_validation(
    'B5',
    {
        validate => 'date',
        criteria => 'between',
        minimum  => '2008-01-01T',
        maximum  => '2008-12-12T',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: date auto, between";
$expected = '<dataValidations count="1"><dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>39448</formula1><formula2>39794</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 34 Time = 0.5 (12:00:00)
#
$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5:B5',
    {
        validate => 'time',
        criteria => '==',
        value    => 0.5,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: time";
$expected = '<dataValidations count="1"><dataValidation type="time" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>0.5</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 35 Time = T12:00:00
#
$worksheet = _new_worksheet( \$got );
$worksheet->{_date_1904} = 0;

$worksheet->data_validation(
    'B5',
    {
        validate => 'time',
        criteria => '==',
        value    => 'T12:00:00',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: time auto";
$expected = '<dataValidations count="1"><dataValidation type="time" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>0.5</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test 36 Custom == 10.
#
$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate => 'Custom',
        criteria => '==',
        value    => 10,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: custom";
$expected = '<dataValidations count="1"><dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>10</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 37 Check the row/col processing: single A1 style cell.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 10,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: range options";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 38 Check the row/col processing: single A1 style range.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5:B10',
    {
        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 10,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: range options";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B10"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 39 Check the row/col processing: single (row, col) style cell.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    4, 1,
    {
        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 10,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: range options";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 40 Check the row/col processing: single (row, col) style range.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    4, 1, 9, 1,
    {
        validate => 'integer',
        criteria => 'between',
        minimum  => 1,
        maximum  => 10,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: range options";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B10"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 41 Check the row/col processing: multiple (row, col) style cells.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    4, 1,
    {
        validate    => 'integer',
        criteria    => 'between',
        minimum     => 1,
        maximum     => 10,
        other_cells => [ [ 4, 3, 4, 3 ] ],
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: range options";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5 D5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 42 Check the row/col processing: multiple (row, col) style cells.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    4, 1,
    {
        validate    => 'integer',
        criteria    => 'between',
        minimum     => 1,
        maximum     => 10,
        other_cells => [ [ 6, 1, 6, 1 ], [ 8, 1, 8, 1 ] ],
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: range options";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5 B7 B9"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 43 Check the row/col processing: multiple (row, col) style cells.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    4, 1, 8, 1,
    {
        validate    => 'integer',
        criteria    => 'between',
        minimum     => 1,
        maximum     => 10,
        other_cells => [ [ 3, 3, 3, 3 ] ],
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: range options";
$expected = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5:B9 D4"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 44 Multiple validations.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate => 'integer',
        criteria => '>',
        value    => 10,
    }
);

$worksheet->data_validation(
    'C10',
    {
        validate => 'integer',
        criteria => '<',
        value    => 10,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: multiple validations";
$expected = '<dataValidations count="2"><dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>10</formula1></dataValidation><dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C10"><formula1>10</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 45 'any' with an input message should produce a dataValidation record.
#

$worksheet = _new_worksheet( \$got );

$worksheet->data_validation(
    'B5',
    {
        validate      => 'any',
        input_title   => 'Input title January',
        input_message => 'Input message February',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: input message only";
$expected = '<dataValidations count="1"><dataValidation allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Input title January" prompt="Input message February" sqref="B5"/></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 46 length validation.
#

$worksheet = _new_worksheet(\$got);

$worksheet->data_validation(
    'A1',
    {
        validate => 'length',
        criteria => 'between',
        minimum  => 5,
        maximum  => 10,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="textLength" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>5</formula1><formula2>10</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 47 length validation.
#

$worksheet = _new_worksheet(\$got);

$worksheet->data_validation(
    'A1',
    {
        validate => 'length',
        criteria => '>',
        value    => 5,
    }
);

$worksheet->_write_data_validations();

$caption  = " \tWorksheet: _write_sheet_views()";
$expected = '<dataValidations count="1"><dataValidation type="textLength" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>5</formula1></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test 48 Date between ranges with formula.
#
$worksheet = _new_worksheet( \$got );
$worksheet->{_date_1904} = 0;

$worksheet->data_validation(
    'B5',
    {
        validate => 'date',
        criteria => 'between',
        minimum  => '2018-01-01T',
        maximum  => '=TODAY()',
    }
);

$worksheet->_write_data_validations();

$caption  = " \tData validation api: date auto, between";
$expected = '<dataValidations count="1"><dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B5"><formula1>43101</formula1><formula2>TODAY()</formula2></dataValidation></dataValidations>';

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );
_is_deep_diff( $got, $expected, $caption );


__END__
