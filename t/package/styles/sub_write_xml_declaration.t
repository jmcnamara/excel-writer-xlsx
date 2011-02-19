###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# reverse('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_style';
use strict;
use warnings;

use Test::More tests => 1;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $style;


###############################################################################
#
# Test the _write_xml_declaration() method.
#
$caption  = " \tStyles: _write_xml_declaration()";
$expected = qq(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n);

$style = _new_style(\$got);

$style->_write_xml_declaration();

is( $got, $expected, $caption );

__END__


