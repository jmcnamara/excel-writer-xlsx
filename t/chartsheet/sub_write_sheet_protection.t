###############################################################################
#
# Tests for Excel::Writer::XLSX::Chartsheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_object';
use strict;
use warnings;

use Test::More tests => 7;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $chartsheet;
my $password;
my %options;


###############################################################################
#
# 1. Test the _write_sheet_protection() method.
#
$caption  = " \tChartsheet: _write_sheet_protection()";
$expected = '<sheetProtection content="1" objects="1"/>';

$password = '';
%options  = ();

$chartsheet = _new_object( \$got, 'Excel::Writer::XLSX::Chartsheet' );

$chartsheet->protect( $password, \%options );
$chartsheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_sheet_protection() method.
#
$caption  = " \tChartsheet: _write_sheet_protection()";
$expected = '<sheetProtection password="83AF" content="1" objects="1"/>';

$password = 'password';
%options  = ();

$chartsheet = _new_object( \$got, 'Excel::Writer::XLSX::Chartsheet' );

$chartsheet->protect( $password, \%options );
$chartsheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_sheet_protection() method.
#
$caption  = " \tChartsheet: _write_sheet_protection()";
$expected = '<sheetProtection content="1"/>';

$password = '';
%options  = ( objects => 0 );

$chartsheet = _new_object( \$got, 'Excel::Writer::XLSX::Chartsheet' );

$chartsheet->protect( $password, \%options );
$chartsheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_sheet_protection() method.
#
$caption  = " \tChartsheet: _write_sheet_protection()";
$expected = '<sheetProtection objects="1"/>';

$password = '';
%options  = ( content => 0);

$chartsheet = _new_object( \$got, 'Excel::Writer::XLSX::Chartsheet' );

$chartsheet->protect( $password, \%options );
$chartsheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_sheet_protection() method.
#
$caption  = " \tChartsheet: _write_sheet_protection()";
$expected = '';

$password = '';
%options  = ( content => 0, objects => 0 );

$chartsheet = _new_object( \$got, 'Excel::Writer::XLSX::Chartsheet' );

$chartsheet->protect( $password, \%options );
$chartsheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_sheet_protection() method.
#
$caption  = " \tChartsheet: _write_sheet_protection()";
$expected = '<sheetProtection password="83AF"/>';

$password = 'password';
%options  = ( content => 0, objects => 0 );

$chartsheet = _new_object( \$got, 'Excel::Writer::XLSX::Chartsheet' );

$chartsheet->protect( $password, \%options );
$chartsheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 7. Test the _write_sheet_protection() method.
#
$caption  = " \tChartsheet: _write_sheet_protection()";
$expected = '<sheetProtection password="83AF" content="1" objects="1"/>';

$password = 'password';
%options = (
    objects               => 1,
    scenarios             => 1,
    format_cells          => 1,
    format_columns        => 1,
    format_rows           => 1,
    insert_columns        => 1,
    insert_rows           => 1,
    insert_hyperlinks     => 1,
    delete_columns        => 1,
    delete_rows           => 1,
    select_locked_cells   => 0,
    sort                  => 1,
    autofilter            => 1,
    pivot_tables          => 1,
    select_unlocked_cells => 0,
);

$chartsheet = _new_object( \$got, 'Excel::Writer::XLSX::Chartsheet' );

$chartsheet->protect( $password, \%options );
$chartsheet->_write_sheet_protection();

is( $got, $expected, $caption );


__END__
