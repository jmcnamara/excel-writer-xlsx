###############################################################################
#
# Tests for Excel::XLSX::Writer::Package::Styles methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::XLSX::Writer::Package::Styles;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $obj     = Excel::XLSX::Writer::Package::Styles->new();
my $writer  = new XML::Writer( OUTPUT => $got_fh );

$obj->{_writer} = $writer;

###############################################################################
#
# Test the _write_cell_style() method.
#
$caption  = " \tStyles: _write_cell_style()";
$expected = '<cellStyle name="Normal" xfId="0" builtinId="0" />';

$obj->_write_cell_style();

is( $got, $expected, $caption );

__END__


