###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# reverse('�'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_workbook';
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
my $workbook;


###############################################################################
#
# Test the _write_xml_declaration() method.
#
$caption  = " \tWorkbook: _write_xml_declaration()";
$expected = qq(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n);

$workbook = _new_workbook(\$got);

$workbook->_write_xml_declaration();

is( $got, $expected, $caption );

__END__


