################################################################################
#
# Tests for Excel::XLSX::Writer methods.
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

my $workbook = Excel::XLSX::Writer->new( $tmp_fh );
my $writer = new XML::Writer( OUTPUT => $got_fh );

$workbook->{_writer} = $writer;

################################################################################
#
# Test the _write_workbook_pr() method.
#
$caption  = " \tWorkbook: _write_workbook_pr()";
$expected = '<workbookPr date1904="1" showInkAnnotation="0" autoCompressPictures="0" />';

$workbook->_write_workbook_pr();

is( $got, $expected, $caption );

__END__


