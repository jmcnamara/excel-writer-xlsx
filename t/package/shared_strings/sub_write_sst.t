###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::SharedStrings methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Package::SharedStrings;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $obj     = Excel::Writer::XLSX::Package::SharedStrings->new();
my $writer  = new XML::Writer( OUTPUT => $got_fh );

$obj->{_writer} = $writer;

###############################################################################
#
# Test the _write_sst() method.
#
$caption  = " \tSharedStrings: _write_sst()";
$expected = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">';

$obj->_write_sst( 7, 3 );

is( $got, $expected, $caption );

__END__


