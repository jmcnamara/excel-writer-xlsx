###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::ContentTypes methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Package::ContentTypes;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $obj     = Excel::Writer::XLSX::Package::ContentTypes->new();
my $writer  = new XML::Writer( OUTPUT => $got_fh );

$obj->{_writer} = $writer;

###############################################################################
#
# Test the _write_override() method.
#
$caption  = " \tContentTypes: _write_override()";
$expected = '<Override PartName="/docProps/core.xml" ContentType="app..." />';

$obj->_write_override( '/docProps/core.xml', 'app...' );

is( $got, $expected, $caption );

__END__


