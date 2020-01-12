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

use Test::More tests => 21;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;
my $filter;
my @matches;

###############################################################################
#
# Test the _write_auto_filter() method with no filter.
#
$expected = '<autoFilter ref="A1:D51"/>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column()" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x == East';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x == East or  x == North';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/><filter val="North"/></filters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x == East and x == North';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters and="1"><customFilter val="East"/><customFilter val="North"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x != East';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="East"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x == S*'; # Begins with character.

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter val="S*"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x != S*'; # Doesn't begin with character.

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="S*"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x == *h'; # Ends with character.

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter val="*h"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x != *h'; # Doesn't end with character.

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="*h"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x =~ *o*'; # Contains character.

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter val="*o*"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x !~ *r*'; # Doesn't contain character.

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val="*r*"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'A', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x == 1000';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><filters><filter val="1000"/></filters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'C', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x != 2000';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="notEqual" val="2000"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'C', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x > 3000';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="greaterThan" val="3000"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'C', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x >= 4000';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="greaterThanOrEqual" val="4000"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'C', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x < 5000';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="lessThan" val="5000"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'C', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x <= 6000';

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters><customFilter operator="lessThanOrEqual" val="6000"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'C', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter() method for the following filter:
#
$filter = 'x >= 1000 and x <= 2000'; # Between.

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="2"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="1000"/><customFilter operator="lessThanOrEqual" val="2000"/></customFilters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column( 'C', $filter );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '$filter' )" );


###############################################################################
#
# Test the _write_auto_filter_list() method for the following filter:
#
@matches = qw( East );

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column_list( 'A', @matches );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '@matches' )" );


###############################################################################
#
# Test the _write_auto_filter_list() method for the following filter:
#
@matches = qw( East North );

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="0"><filters><filter val="East"/><filter val="North"/></filters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column_list( 'A', @matches );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '@matches' )" );


###############################################################################
#
# Test the _write_auto_filter_list() method for the following filter:
#
@matches = qw( February January July June );

$expected = '<autoFilter ref="A1:D51"><filterColumn colId="3"><filters><filter val="February"/><filter val="January"/><filter val="July"/><filter val="June"/></filters></filterColumn></autoFilter>';

$worksheet = _new_worksheet( \$got );
$worksheet->{_name} = 'Sheet1';
$worksheet->autofilter( 'A1:D51' );

$worksheet->filter_column_list( 'D', @matches );
$worksheet->_write_auto_filter();

is( $got, $expected, " \tWorksheet: filter_column( '@matches' )" );



__END__


