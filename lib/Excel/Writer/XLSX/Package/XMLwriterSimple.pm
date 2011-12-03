package Excel::Writer::XLSX::Package::XMLwriterSimple;

###############################################################################
#
# XMLwriterSimple - Light weight re-implementation of XML::Writer.
#
# Used in conjunction with Excel::Writer::XLSX.
#
# Copyright 2000-2011, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.008002;
use strict;
use warnings;
use Exporter;
use Carp;

our @ISA     = qw(Exporter);
our $VERSION = '0.39';

#
# NOTE: this module is a light weight re-implementation of XML::Writer. See
# the Pod docs below for a full explanation. The methods implemented below
# are the main XML::Writer methods used by Excel::Writer::XLSX.
# See XML::Writer for more detailed information on these methods.
#

###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;

    my $self = { _fh => shift };

    bless $self, $class;

    return $self;
}


###############################################################################
#
# xmlDecl()
#
# Write the XML declaration at the start of an XML document.
#
sub xmlDecl {

    my $self = shift;

    print { $self->{_fh} }
      qq(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n);

}


###############################################################################
#
# startTag()
#
# Write an XML start tag with optional attributes.
#
sub startTag {

    my $self       = shift;
    my $tag        = shift;
    my @attributes = @_;

    print { $self->{_fh} } "<$tag";

    while ( @attributes ) {
        my $key   = shift @attributes;
        my $value = shift @attributes;
        $value = _escape_xml_chars( $value );

        print { $self->{_fh} } qq( $key="$value");
    }

    print { $self->{_fh} } ">";
}


###############################################################################
#
# endTag()
#
# Write an XML end tag.
#
sub endTag {

    my $self = shift;
    my $tag  = shift;

    print { $self->{_fh} } "</$tag>";
}


###############################################################################
#
# emptyTag()
#
# Write an empty XML tag with optional attributes.
#
sub emptyTag {

    my $self       = shift;
    my $tag        = shift;
    my @attributes = @_;

    print { $self->{_fh} } "<$tag";

    while ( @attributes ) {
        my $key   = shift @attributes;
        my $value = shift @attributes;
        $value = _escape_xml_chars( $value );

        print { $self->{_fh} } qq( $key="$value");
    }

    # Note extra space before closing tag like XML::Writer.
    print { $self->{_fh} } " />";

}


###############################################################################
#
# dataElement()
#
# Write an XML element containing data with optional attributes.
# XML characters in the data are escaped.
#
sub dataElement {

    my $self       = shift;
    my $tag        = shift;
    my $data       = shift;
    my @attributes = @_;

    print { $self->{_fh} } "<$tag";

    while ( @attributes ) {
        my $key   = shift @attributes;
        my $value = shift @attributes;
        $value = _escape_xml_chars( $value );

        print { $self->{_fh} } qq( $key="$value");
    }

    $data = _escape_xml_chars( $data );

    print { $self->{_fh} } ">";
    print { $self->{_fh} } $data;
    print { $self->{_fh} } "</$tag>";
}


###############################################################################
#
# characters()
#
# For compatibility with XML::Writer only.
#
sub characters {

    my $self = shift;
    my $data = shift;

    $data = _escape_xml_chars( $data );

    print { $self->{_fh} } $data;
}


###############################################################################
#
# end()
#
# For compatibility with XML::Writer only.
#
sub end {

    my $self = shift;

    print { $self->{_fh} } "\n";
}


###############################################################################
#
# getOutput()
#
# Return the output filehandle.
#
sub getOutput {

    my $self = shift;

    return $self->{_fh};
}


###############################################################################
#
# _escape_xml_chars()
#
# Escape XML characters.
#
sub _escape_xml_chars {

    my $str = defined $_[0] ? $_[0] : '';

    return $str if $str !~ m/[&<>"]/;

    for ( $str ) {
        s/&/&amp;/g;
        s/</&lt;/g;
        s/>/&gt;/g;
        s/"/&quot;/g;
    }

    return $str;
}


1;


__END__

=pod

=head1 NAME

XMLwriterSimple - Light weight re-implementation of XML::Writer.

This module is used internally by L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used by L<Excel::Writer::XLSX> for writing XML documents. It is a light weight re-implementation of L<XML::Writer>.

XMLwriterSimple is approximately twice as fast as L<XML::Writer>. This speed is achieved at the expense of error and correctness checking. In addition not all of the L<XML::Writer> methods are implemented. As such, XMLwriterSimple is not recommended for use outside of Excel::Writer::XLSX.

If required XMLwriterSimple can be overridden and XML::Writer can be used in its place by setting an C<_EXCEL_WRITER_XLSX_USE_XML_WRITER> environmental variable:

    export _EXCEL_WRITER_XLSX_USE_XML_WRITER=1
    perl example.pl

    Or for one off programs:

    _EXCEL_WRITER_XLSX_USE_XML_WRITER=1 perl example.pl

This technique is used for verifying the test suite with both XMLwriterSimple and XML::Writer:

    _EXCEL_WRITER_XLSX_USE_XML_WRITER=1 prove -l -r t

=head1 SEE ALSO

L<XML::Writer>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMXI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Writer::XLSX>.

=cut
