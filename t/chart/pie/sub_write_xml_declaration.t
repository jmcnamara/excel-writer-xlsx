###############################################################################
#
# Tests for Excel::Writer::XLSX::Chart methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions qw(_new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Chart::Pie;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $obj = _new_object( \$got, 'Excel::Writer::XLSX::Chart::Pie' );


###############################################################################
#
# Test the xml_declaration() method.
#
$caption  = " \tApp: xml_declaration()";
$expected = qq(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n);

$obj->xml_declaration();

is( $got, $expected, $caption );

__END__


