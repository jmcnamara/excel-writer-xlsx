###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Package::Styles;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $obj     = Excel::Writer::XLSX::Package::Styles->new();
my $writer  = new XML::Writer( OUTPUT => $got_fh );

$obj->{_writer} = $writer;

###############################################################################
#
# Test the _write_style_sheet() method.
#
$caption  = " \tStyles: _write_style_sheet()";
$expected = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

$obj->_write_style_sheet();

is( $got, $expected, $caption );

__END__


