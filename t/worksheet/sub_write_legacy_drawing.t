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

use Test::More tests => 1;


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
# Test the _write_legacy_drawing() method.
#
$caption  = " \tWorksheet: _write_legacy_drawing()";
$expected = '<legacyDrawing r:id="rId1"/>';

$worksheet = _new_worksheet(\$got);

$worksheet->{_has_vml} = 1;

$worksheet->_write_legacy_drawing();

is( $got, $expected, $caption );

__END__


