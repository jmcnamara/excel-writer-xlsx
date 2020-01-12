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


###############################################################################
#
# Test the _write_hyperlink_external() method.
#
$caption  = " \tWorksheet: _write_hyperlink_external()";
$expected = '<hyperlink ref="A1" r:id="rId1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_hyperlink_external( 0, 0, 1 );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_hyperlink_internal() method.
#
$caption  = " \tWorksheet: _write_hyperlink_internal()";
$expected = '<hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_hyperlink_internal( 0, 0, 'Sheet2!A1', 'Sheet2!A1' );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_hyperlink_internal() method.
#
$caption  = " \tWorksheet: _write_hyperlink_internal()";
$expected = q(<hyperlink ref="A5" location="'Data Sheet'!D5" display="'Data Sheet'!D5"/>);

$worksheet = _new_worksheet(\$got);

$worksheet->_write_hyperlink_internal( 4, 0, "'Data Sheet'!D5", "'Data Sheet'!D5" );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_hyperlink_internal() method.
#
$caption  = " \tWorksheet: _write_hyperlink_internal()";
$expected = '<hyperlink ref="A18" location="Sheet2!A1" tooltip="Screen Tip 1" display="Sheet2!A1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->_write_hyperlink_internal( 17, 0, 'Sheet2!A1', 'Sheet2!A1', 'Screen Tip 1' );

is( $got, $expected, $caption );


__END__


