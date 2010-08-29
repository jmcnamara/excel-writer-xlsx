################################################################################
#
# Tests for Excel::XLSX::Writer::Worksheet methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::XLSX::Writer;
use XML::Writer;

use Test::More tests => 1;

################################################################################
#
# Tests setup.
#
my $expected;
my $caption;

open my $tmp_fh, '>', \my $tmp or die "Failed to open filehandle: $!";
open my $got_fh, '>', \my $got or die "Failed to open filehandle: $!";

my $workbook  = Excel::XLSX::Writer->new( $tmp_fh );
my $worksheet = $workbook->add_worksheet;
my $writer    = new XML::Writer( OUTPUT => $got_fh );

$worksheet->{_writer} = $writer;

################################################################################
#
# Test the _write_page_setup() method.
#
$caption  = " \tWorksheet: _write_page_setup()";
$expected = '<pageSetup paperSize="0" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292" />';

$worksheet->_write_page_setup();

is( $got, $expected, $caption );

__END__


