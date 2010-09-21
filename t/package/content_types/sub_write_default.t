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
# Test the _write_default() method.
#
$caption  = " \tContentTypes: _write_default()";
$expected = '<Default Extension="xml" ContentType="application/xml" />';

$obj->_write_default( 'xml', 'application/xml' );

is( $got, $expected, $caption );

__END__


