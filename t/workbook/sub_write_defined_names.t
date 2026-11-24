###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# Copyright 2000-2026, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions qw(_new_workbook _is_deep_diff _got_to_aref);
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
my $workbook;


###############################################################################
#
# Test the _write_defined_names() method.
#
$caption  = " \tWorkbook: _write_defined_names()";
$expected = '<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName></definedNames>';

$workbook = _new_workbook(\$got);
$workbook->{_defined_names} = [ [ '_xlnm.Print_Titles', 0, 'Sheet1!$1:$1' ] ];

$workbook->_write_defined_names();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_defined_names() method.
#
$caption  = " \tWorkbook: _write_defined_names()";
$expected = q(<definedNames><definedName name="_Egg">Sheet1!$A$1</definedName><definedName name="_Fog">Sheet1!$A$1</definedName><definedName name="aaa" localSheetId="1">Sheet2!$A$1</definedName><definedName name="Abc">Sheet1!$A$1</definedName><definedName name="Bar" localSheetId="2">'Sheet 3'!$A$1</definedName><definedName name="Bar" localSheetId="0">Sheet1!$A$1</definedName><definedName name="Bar" localSheetId="1">Sheet2!$A$1</definedName><definedName name="Baz">0.98</definedName><definedName name="car" localSheetId="2">"Saab 900"</definedName></definedNames>);

$workbook = _new_workbook(\$got);

$workbook->add_worksheet();
$workbook->add_worksheet();
$workbook->add_worksheet('Sheet 3');

$workbook->define_name( q('Sheet 3'!Bar), q(='Sheet 3'!$A$1) );
$workbook->define_name( q(Abc),           q(=Sheet1!$A$1) );
$workbook->define_name( q(Baz),           q(=0.98) );
$workbook->define_name( q(Sheet1!Bar),    q(=Sheet1!$A$1) );
$workbook->define_name( q(Sheet2!Bar),    q(=Sheet2!$A$1) );
$workbook->define_name( q(Sheet2!aaa),    q(=Sheet2!$A$1) );
$workbook->define_name( q('Sheet 3'!car), q(="Saab 900") );
$workbook->define_name( q(_Egg),          q(=Sheet1!$A$1) );
$workbook->define_name( q(_Fog),          q(=Sheet1!$A$1) );

$workbook->_prepare_defined_names();
$workbook->_write_defined_names();

$expected = _got_to_aref( $expected );
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test define_name() with emojis and Excel's 255 UTF-16 unit limit. Characters
# above U+FFFF, such as emojis, count as 2.
#
$caption  = " \tWorkbook: define_name() UTF-16 length";
$workbook = _new_workbook(\$got);

{
    local $SIG{__WARN__} = sub { };

    my $result = $workbook->define_name( "\x{1F600}" x 127 . 'a', '=1' );
    isnt( $result, -1, $caption );

    $result = $workbook->define_name( "\x{1F600}" x 128, '=1' );
    is( $result, -1, $caption );
}


__END__
