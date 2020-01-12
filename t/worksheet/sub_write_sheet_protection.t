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

use Test::More tests => 18;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;
my $password;
my %options;


###############################################################################
#
# 1. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1"/>';

$password = '';
%options  = ();

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection password="83AF" sheet="1" objects="1" scenarios="1"/>';

$password = 'password';
%options  = ();

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="1"/>';

$password = '';
%options  = ( select_locked_cells => 0 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0"/>';

$password = '';
%options  = ( format_cells => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0"/>';

$password = '';
%options  = ( format_columns => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatRows="0"/>';

$password = '';
%options  = (  format_rows => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 7. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" insertColumns="0"/>';

$password = '';
%options  = ( insert_columns => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 8. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" insertRows="0"/>';

$password = '';
%options  = ( insert_rows => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 9. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" insertHyperlinks="0"/>';

$password = '';
%options  = ( insert_hyperlinks => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 10. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" deleteColumns="0"/>';

$password = '';
%options  = ( delete_columns => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 11. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" deleteRows="0"/>';

$password = '';
%options  = ( delete_rows => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 12. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" sort="0"/>';

$password = '';
%options  = ( sort => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 13. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" autoFilter="0"/>';

$password = '';
%options  = ( autofilter => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 14. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" pivotTables="0"/>';

$password = '';
%options  = ( pivot_tables => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 15. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" scenarios="1"/>';

$password = '';
%options  = ( objects => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 16. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1"/>';

$password = '';
%options  = ( scenarios => 1 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 17. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" selectLockedCells="1" selectUnlockedCells="1"/>';

$password = '';
%options  = ( format_cells => 1, select_locked_cells => 0, select_unlocked_cells => 0 );

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


###############################################################################
#
# 18. Test the _write_sheet_protection() method.
#
$caption  = " \tWorksheet: _write_sheet_protection()";
$expected = '<sheetProtection password="996B" sheet="1" formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" selectLockedCells="1" sort="0" autoFilter="0" pivotTables="0" selectUnlockedCells="1"/>';

$password = 'drowssap';
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

$worksheet = _new_worksheet(\$got);

$worksheet->protect( $password, \%options );
$worksheet->_write_sheet_protection();

is( $got, $expected, $caption );


__END__


