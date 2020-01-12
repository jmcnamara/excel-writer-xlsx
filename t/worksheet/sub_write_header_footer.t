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

use Test::More tests => 5;


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
# Test the _write_odd_header() method.
#
$caption  = " \tWorksheet: _write_odd_header()";
$expected = '<oddHeader>Page &amp;P of &amp;N</oddHeader>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_header('Page &P of &N');

$worksheet->_write_odd_header();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_odd_footer() method.
#
$caption  = " \tWorksheet: _write_odd_footer()";
$expected = '<oddFooter>&amp;F</oddFooter>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_footer('&F');

$worksheet->_write_odd_footer();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_header_footer() method. Header only.
#
$caption  = " \tWorksheet: _write_header_footer()";
$expected = '<headerFooter><oddHeader>Page &amp;P of &amp;N</oddHeader></headerFooter>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_header('Page &P of &N');

$worksheet->_write_header_footer();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_header_footer() method. Footer only.
#
$caption  = " \tWorksheet: _write_header_footer()";
$expected = '<headerFooter><oddFooter>&amp;F</oddFooter></headerFooter>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_footer('&F');

$worksheet->_write_header_footer();

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_header_footer() method. Header and footer.
#
$caption  = " \tWorksheet: _write_header_footer()";
$expected = '<headerFooter><oddHeader>Page &amp;P of &amp;N</oddHeader><oddFooter>&amp;F</oddFooter></headerFooter>';

$worksheet = _new_worksheet(\$got);
$worksheet->set_header('Page &P of &N');
$worksheet->set_footer('&F');

$worksheet->_write_header_footer();

is( $got, $expected, $caption );


__END__


