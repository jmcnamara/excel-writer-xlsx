package TestFunctions;

###############################################################################
#
# TestFunctions - Helper functions for Excel::Writer::XLSX test cases.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use 5.010000;
use Exporter;
use strict;
use warnings;
use Test::More;
use Excel::Writer::XLSX;
use XML::Writer;


our @ISA         = qw(Exporter);
our @EXPORT      = ();
our %EXPORT_TAGS = ();
our @EXPORT_OK   = qw(
  _expected_to_aref
  _got_to_aref
  _is_deep_diff
  _new_worksheet
  _new_workbook
  _new_style
);

our $VERSION = '0.01';


###############################################################################
#
# Turn the embedded XML in the __DATA__ section of the calling test program
# into an array ref for comparison testing. Also performs some minor string
# formatting to make comparison easier with _got_to_aref().
#
# The XML data in the testcases is taken from Excel 2007 files with formatting
# via "xmllint --format".
#
sub _expected_to_aref {

    my @data;

    while ( <main::DATA> ) {
        chomp;
        next unless /\S/;    # Skip blank lines.
        s{/>$}{ />};         # Add space before element end like XML::Writer.
        s{^\s+}{};           # Remove leading whitespace from XML.
        push @data, $_;
    }

    return \@data;
}


###############################################################################
#
# Convert an XML string returned by the XML::Writer subclasses into an
# array ref for comparison testing with _expected_to_aref().
#
sub _got_to_aref {

    my $xml_str = shift;

    $xml_str =~ s/\n//g;

    # Split the XML into chunks at element boundaries.
    my @data = split /(?<=>)(?=<)/, $xml_str;

    return \@data;
}


###############################################################################
#
# Use Test::Differences::eq_or_diff() where available or else fall back to
# using Test::More::is_deeply().
#
sub _is_deep_diff {
    my ( $got, $expected, $caption, ) = @_;

    eval {
        require Test::Differences;
        Test::Differences->import();
    };

    if ( !$@ ) {
        eq_or_diff( $got, $expected, $caption, { context => 1 } );
    }
    else {
        is_deeply( $got, $expected, $caption );
    }

}


###############################################################################
#
# Create a new XML::Writer sub-classed object based on a class name and bind
# the output to the supplied scalar ref for testing. Calls to the objects XML
# writing subs will add the output to the scalar.
#
sub _new_object {

    my $got_ref = shift;
    my $class   = shift;

    open my $got_fh, '>', $got_ref or die "Failed to open filehandle: $!";

    my $object = new $class;
    my $writer = new XML::Writer( OUTPUT => $got_fh );

    $object->{_writer} = $writer;

    return $object;
}


###############################################################################
#
# Create a new Worksheet object and bind the output to the supplied scalar ref.
#
sub _new_worksheet {

    my $got_ref = shift;

    return _new_object( $got_ref, 'Excel::Writer::XLSX::Worksheet' );
}


###############################################################################
#
# Create a new Style object and bind the output to the supplied scalar ref.
#
sub _new_style {

    my $got_ref = shift;

    return _new_object( $got_ref, 'Excel::Writer::XLSX::Package::Styles' );
}


###############################################################################
#
# Create a new Workbook object and bind the output to the supplied scalar ref.
# This is slightly different than the previous cases since the constructor
# requires a filename/filehandle.
#
sub _new_workbook {

    my $got_ref = shift;

    open my $got_fh, '>', $got_ref or die "Failed to open filehandle: $!";
    open my $tmp_fh, '>', \my $tmp or die "Failed to open filehandle: $!";

    my $workbook = Excel::Writer::XLSX->new( $tmp_fh );
    my $writer = new XML::Writer( OUTPUT => $got_fh );

    $workbook->{_writer} = $writer;

    return $workbook;
}


1;


__END__

