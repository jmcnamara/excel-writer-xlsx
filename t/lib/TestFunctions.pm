package TestFunctions;

###############################################################################
#
# TestFunctions - Helper functions for Excel::Writer::XLSX test cases.
#
# reverse ('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
#

use 5.008002;
use Exporter;
use strict;
use warnings;
use Test::More;
use Excel::Writer::XLSX;


our @ISA         = qw(Exporter);
our @EXPORT      = ();
our %EXPORT_TAGS = ();
our @EXPORT_OK   = qw(
  _expected_to_aref
  _expected_vml_to_aref
  _got_to_aref
  _is_deep_diff
  _new_object
  _new_worksheet
  _new_workbook
  _new_style
  _compare_xlsx_files
);

our $VERSION = '0.05';


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

    # Ignore warning for files that don't have a 'main::DATA'.
    no warnings 'once';

    while ( <main::DATA> ) {
        chomp;
        next unless /\S/;    # Skip blank lines.
        s{^\s+}{};           # Remove leading whitespace from XML.
        push @data, $_;
    }

    return \@data;
}


###############################################################################
#
# Turn the embedded VML in the __DATA__ section of the calling test program
# into an array ref for comparison testing.
#
sub _expected_vml_to_aref {

    # Ignore warning for files that don't have a 'main::DATA'.
    no warnings 'once';

    my $vml_str = do { local $/; <main::DATA> };

    my @vml = _vml_str_to_array( $vml_str );

    return \@vml;
}


###############################################################################
#
# Convert an XML string returned by the XMLWriter subclasses into an
# array ref for comparison testing with _expected_to_aref().
#
sub _got_to_aref {

    my $xml_str = shift;

    # Remove the newlines after the XML declaration and any others.
    $xml_str =~ s/[\r\n]//g;

    # Split the XML into chunks at element boundaries.
    my @data = split /(?<=>)(?=<)/, $xml_str;

    return \@data;
}

###############################################################################
#
# _xml_str_to_array()
#
# Convert an XML string into an array for comparison testing.
#
sub _xml_str_to_array {

    my $xml_str = shift;
    my @xml     = @{ _got_to_aref( $xml_str ) };

    #s{ />$}{/>} for @xml;

    return @xml;
}

###############################################################################
#
# _vml_str_to_array()
#
# Convert an Excel generated VML string into an array for comparison testing.
#
# The VML data in the testcases is taken from Excel 2007 files. The data has
# to be massaged significantly to make it suitable for comparison.
#
# Excel::Writer::XLSX produced VML can be parsed as ordinary XML.
#
sub _vml_str_to_array {

    my $vml_str = shift;
    my @vml = split /[\r\n]+/, $vml_str;

    $vml_str = '';

    for ( @vml ) {

        chomp;
        next unless /\S/;    # Skip blank lines.

        s/^\s+//;            # Remove leading whitespace.
        s/\s+$//;            # Remove trailing whitespace.
        s/\'/"/g;            # Convert VMLs attribute quotes.

        $_ .= " "  if /"$/;  # Add space between attributes.
        $_ .= "\n" if />$/;  # Add newline after element end.

        s/></>\n</g;         # Split multiple elements.

        chomp if $_ eq "<x:Anchor>\n";    # Put all of Anchor on one line.

        $vml_str .= $_;
    }

    return ( split "\n", $vml_str );
}


###############################################################################
#
# _compare_xlsx_files()
#
# Compare two XLSX files by extracting the XML files from each archive and
# comparing them.
#
# This is used to compare an "expected" file produced by Excel with a "got"
# file produced by Excel::Writer::XLSX.
#
# In order to compare the XLSX files we convert the data in each XML file.
# contained in the zip archive into arrays of XML elements to make identifying
# differences easier.
#
# This function returns 3 elements suitable for _is_deep_diff() comparison:
#    return ( $got_aref, $expected_aref, $caption)
#
sub _compare_xlsx_files {

    my $got_filename    = shift;
    my $exp_filename    = shift;
    my $ignore_members  = shift;
    my $ignore_elements = shift;
    my $got_zip         = Archive::Zip->new();
    my $exp_zip         = Archive::Zip->new();
    my @got_xml;
    my @exp_xml;

    # Suppress Archive::Zip error reporting. We will handle errors.
    Archive::Zip::setErrorHandler( sub { } );

    # Test the $got file exists.
    if ( $got_zip->read( $got_filename ) != 0 ) {
        my $error = "Excel::Write::XML generated file not found.";
        return ( [$error], [$got_filename], " _compare_xlsx_files(). Files." );
    }

    # Test the $exp file exists.
    if ( $exp_zip->read( $exp_filename ) != 0 ) {
        my $error = "Excel generated comparison file not found.";
        return ( [$error], [$exp_filename], " _compare_xlsx_files(). Files." );
    }

    # The zip "members" are the files in the XLSX container.
    my @got_members = sort $got_zip->memberNames();
    my @exp_members = sort $exp_zip->memberNames();

    # Ignore some test specific filenames.
    if ( defined $ignore_members && @$ignore_members ) {
        my $ignore_regex = join '|', @$ignore_members;

        @got_members = grep { !/$ignore_regex/ } @got_members;
        @exp_members = grep { !/$ignore_regex/ } @exp_members;
    }

    # Check that each XLSX container has the same file members.
    if ( !_arrays_equal( \@got_members, \@exp_members ) ) {
        return ( \@got_members, \@exp_members,
            ' _compare_xlsx_files(): Members.' );
    }

    # Compare each file in the XLSX containers.
    for my $filename ( @exp_members ) {
        my $got_xml_str = $got_zip->contents( $filename );
        my $exp_xml_str = $exp_zip->contents( $filename );

        # Remove dates and user specific data from the core.xml data.
        if ( $filename eq 'docProps/core.xml' ) {
            $exp_xml_str =~ s/ ?John//g;
            $exp_xml_str =~ s/\d\d\d\d-\d\d-\d\dT\d\d\:\d\d:\d\dZ//g;
            $got_xml_str =~ s/\d\d\d\d-\d\d-\d\dT\d\d\:\d\d:\d\dZ//g;
        }

        # Remove workbookView dimensions which are almost always different.
        if ( $filename eq 'xl/workbook.xml' ) {
            $exp_xml_str =~ s{<workbookView[^>]*>}{<workbookView/>};
            $got_xml_str =~ s{<workbookView[^>]*>}{<workbookView/>};
        }

        # Remove the calcPr elements which may have different Excel version ids.
        if ( $filename eq 'xl/workbook.xml' ) {
            $exp_xml_str =~ s{<calcPr[^>]*>}{<calcPr/>};
            $got_xml_str =~ s{<calcPr[^>]*>}{<calcPr/>};
        }

        # Remove printer specific settings from Worksheet pageSetup elements.
        if ( $filename =~ m(xl/worksheets/sheet\d.xml) ) {
            $exp_xml_str =~ s/horizontalDpi="200" //;
            $exp_xml_str =~ s/verticalDpi="200" //;
            $exp_xml_str =~ s/(<pageSetup[^>]*) r:id="rId1"/$1/;
        }

        # Remove Chart pageMargin dimensions which are almost always different.
        if ( $filename =~ m(xl/charts/chart\d.xml) ) {
            $exp_xml_str =~ s{<c:pageMargins[^>]*>}{<c:pageMargins/>};
            $got_xml_str =~ s{<c:pageMargins[^>]*>}{<c:pageMargins/>};
        }

        if ( $filename =~ /.vml$/ ) {
            @got_xml = _xml_str_to_array( $got_xml_str );
            @exp_xml = _vml_str_to_array( $exp_xml_str );
        }
        else {
            @got_xml = _xml_str_to_array( $got_xml_str );
            @exp_xml = _xml_str_to_array( $exp_xml_str );
        }

        # Ignore test specific XML elements for defined filenames.
        if ( defined $ignore_elements && exists $ignore_elements->{$filename} )
        {
            my @ignore_elements = @{ $ignore_elements->{$filename} };

            if ( @ignore_elements ) {
                my $ignore_regex = join '|', @ignore_elements;
                @got_xml = grep { !/$ignore_regex/ } @got_xml;
                @exp_xml = grep { !/$ignore_regex/ } @exp_xml;
            }
        }

        # Reorder the XML elements in the XLSX relationship files.
        if ( $filename eq '[Content_Types].xml' || $filename =~ /.rels$/ ) {
            @got_xml = _sort_rel_file_data( @got_xml );
            @exp_xml = _sort_rel_file_data( @exp_xml );
        }

        # Comparison of the XML elements in each file.
        if ( !_arrays_equal( \@got_xml, \@exp_xml ) ) {
            return ( \@got_xml, \@exp_xml,
                " _compare_xlsx_files(): $filename" );
        }
    }

    # Files were the same. Return values that will evaluate to a test pass.
    return ( ['ok'], ['ok'], ' _compare_xlsx_files()' );
}


###############################################################################
#
# _arrays_equal()
#
# Compare two array refs for equality.
#
sub _arrays_equal {

    my $exp = shift;
    my $got = shift;

    if ( @$exp != @$got ) {
        return 0;
    }

    for my $i ( 0 .. @$exp - 1 ) {
        if ( $exp->[$i] ne $got->[$i] ) {
            return 0;
        }
    }

    return 1;
}


###############################################################################
#
# _sort_rel_file_data()
#
# Re-order the relationship elements in an array of XLSX XML rel (relationship)
# data. This is necessary for comparison since Excel can produce the elements
# in a semi-random order.
#
sub _sort_rel_file_data {

    my @xml_elements = @_;
    my $header       = shift @xml_elements;
    my $tail         = pop @xml_elements;

    # Sort the relationship elements.
    @xml_elements = sort @xml_elements;

    return $header, @xml_elements, $tail;
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
# Create a new XML writer sub-classed object based on a class name and bind
# the output to the supplied scalar ref for testing. Calls to the objects XML
# writing subs will add the output to the scalar.
#
sub _new_object {

    my $got_ref = shift;
    my $class   = shift;

    open my $got_fh, '>', $got_ref or die "Failed to open filehandle: $!";

    my $object = $class->new( $got_fh );

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

    $workbook->{_fh} = $got_fh;

    return $workbook;
}


1;


__END__

