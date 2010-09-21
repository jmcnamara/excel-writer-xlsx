###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Styles methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Package::Styles;
use Excel::Writer::XLSX::Format;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $obj = Excel::Writer::XLSX::Package::Styles->new();
my $writer = new XML::Writer( OUTPUT => $got_fh );

$obj->{_writer} = $writer;

###############################################################################
#
# Test the _write_fonts() method.
#
$caption = " \tStyles: _write_fonts()";
$expected =
'<fonts count="1"><font><sz val="11" /><color theme="1" /><name val="Calibri" /><family val="2" /><scheme val="minor" /></font></fonts>';


my @formats = ( Excel::Writer::XLSX::Format->new( 0, has_font => 1 ) );
my $num_fonts = 1;

$obj->_set_format_properties( \@formats, $num_fonts );
$obj->_write_fonts();

is( $got, $expected, $caption );

__END__


