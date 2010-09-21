###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Core methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Package::Core;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $obj    = Excel::Writer::XLSX::Package::Core->new();
my $writer = new XML::Writer( OUTPUT => $got_fh );

$obj->{_writer} = $writer;

###############################################################################
#
# Test the _write_xml_declaration() method.
#
$caption  = " \tCore: _write_xml_declaration()";
$expected = qq(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n);

$obj->_write_xml_declaration();

is( $got, $expected, $caption );

__END__


