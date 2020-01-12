###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_style';
use strict;
use warnings;

use Test::More tests => 2;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $style;
my @formats;
my $num_format_count;

###############################################################################
#
# 1. Test the _write_num_fmts() method.
#
$caption  = " \tStyles: _write_num_fmts()";
$expected = undef;

$num_format_count = 0;

$style = _new_style( \$got );

$style->_write_num_fmts();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_num_fmts() method.
#
$num_format_count = 1;
$caption          = " \tStyles: _write_num_fmts()";
$expected         =  '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.0"/></numFmts>';

@formats = (
    Excel::Writer::XLSX::Format->new(
        0,
        {},
        num_format_index => 164,
        num_format       => '#,##0.0'
    )
);

$num_format_count = 1;

$got = ''; # Since it was previously undef.

$style = _new_style( \$got );
$style->_set_style_properties( \@formats, undef,  undef, $num_format_count );

$style->_write_num_fmts();

is( $got, $expected, $caption );

__END__


