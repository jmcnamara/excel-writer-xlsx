###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet data validation methods.
#
# reverse('(c)'), June 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_got_to_aref _is_deep_diff _new_worksheet);
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





__END__
