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

use Test::More tests => 6;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $worksheet;
my $min;
my $max;
my $width;
my $format;
my $hidden;
my $level;
my $collapsed;


###############################################################################
#
# 1. Test the _write_col_info() method.
#
$min       = 1;
$max       = 3;
$width     = 5;
$format    = undef;
$hidden    = 0;
$level     = 0;
$collapsed = 0;

$caption  = " \tWorksheet: _write_col_info( $min, $max )";
$expected = '<col min="2" max="4" width="5.7109375" customWidth="1"/>';

$worksheet = _new_worksheet( \$got );
$worksheet->_write_col_info( $min, $max, $width, $format, $hidden );

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_col_info() method.
#
$min       = 5;
$max       = 5;
$width     = 8;
$format    = undef;
$hidden    = 1;
$level     = 0;
$collapsed = 0;

$caption  = " \tWorksheet: _write_col_info( $min, $max )";
$expected = '<col min="6" max="6" width="8.7109375" hidden="1" customWidth="1"/>';

$worksheet = _new_worksheet( \$got );
$worksheet->_write_col_info( $min, $max, $width, $format, $hidden );

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_col_info() method.
#
$min       = 7;
$max       = 7;
$width     = undef;
$format    = Excel::Writer::XLSX::Format->new( {}, {}, xf_index => 1 );
$hidden    = 0;
$level     = 0;
$collapsed = 0;

$caption  = " \tWorksheet: _write_col_info( $min, $max )";
$expected = '<col min="8" max="8" width="9.140625" style="1"/>';

$worksheet = _new_worksheet( \$got );
$worksheet->_write_col_info( $min, $max, $width, $format, $hidden );

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_col_info() method.
#
$min       = 8;
$max       = 8;
$width     = 8.43;
$format    = Excel::Writer::XLSX::Format->new( {}, {}, xf_index => 1 );
$hidden    = 0;
$level     = 0;
$collapsed = 0;

$caption  = " \tWorksheet: _write_col_info( $min, $max )";
$expected = '<col min="9" max="9" width="9.140625" style="1"/>';

$worksheet = _new_worksheet( \$got );
$worksheet->_write_col_info( $min, $max, $width, $format, $hidden );

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_col_info() method.
#
$min       = 9;
$max       = 9;
$width     = 2;
$format    = undef;
$hidden    = 0;
$level     = 0;
$collapsed = 0;

$caption  = " \tWorksheet: _write_col_info( $min, $max )";
$expected = '<col min="10" max="10" width="2.7109375" customWidth="1"/>';

$worksheet = _new_worksheet( \$got );
$worksheet->_write_col_info( $min, $max, $width, $format, $hidden );

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_col_info() method.
#
$min       = 11;
$max       = 11;
$width     = undef;
$format    = undef;
$hidden    = 1;
$level     = 0;
$collapsed = 0;

$caption  = " \tWorksheet: _write_col_info( $min, $max )";
$expected = '<col min="12" max="12" width="0" hidden="1" customWidth="1"/>';

$worksheet = _new_worksheet( \$got );

$worksheet->_write_col_info( $min, $max, $width, $format, $hidden );

is( $got, $expected, $caption );


__END__


