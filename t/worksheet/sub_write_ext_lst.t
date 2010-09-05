###############################################################################
#
# Tests for Excel::XLSX::Writer::Worksheet methods.
#
# reverse('�'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::XLSX::Writer;
use XML::Writer;

use Test::More tests => 1;

###############################################################################
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

###############################################################################
#
# Test the _write_ext_lst() method.
#
$caption  = " \tWorksheet: _write_ext_lst()";
$expected = '<extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="1" OnePage="0" WScale="0" /></ext></extLst>';

$worksheet->_write_ext_lst();

is( $got, $expected, $caption );

__END__


