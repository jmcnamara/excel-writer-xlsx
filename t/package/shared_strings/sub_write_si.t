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
# Test the _write_si() method.
#
$caption  = " \tSharedStrings: _write_si()";
$expected = '<si><t>neptune</t></si>';

$obj->_write_si('neptune');

is( $got, $expected, $caption );

__END__


