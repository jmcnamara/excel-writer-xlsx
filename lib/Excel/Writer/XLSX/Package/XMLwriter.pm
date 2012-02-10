package Excel::Writer::XLSX::Package::XMLwriter;

###############################################################################
#
# XMLwriter - A base class for the Excel::Writer::XLSX writer classes.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2012, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.008002;
use strict;
use warnings;
use Exporter;
use Carp;
use IO::File;
use Excel::Writer::XLSX::Package::XMLwriterSimple;

our @ISA     = qw(Exporter);
our $VERSION = '0.46';


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;

    my $self = { _writer => undef };

    bless $self, $class;

    return $self;
}


###############################################################################
#
# _set_xml_writer()
#
# Set the XML writer class for the object. For speed we use the internal
# Excel::Writer::XLSX::Package::XMLwriterSimple class but for XML error
# and correctness checking we can use the CPAN module XML::Writer.
#
# In general use we use XMLwriterSimple but maintain compatibility with
# XML::Writer for testing purposes. We can choose between the two using an
# environmental variable:
#
#    export _EXCEL_WRITER_XLSX_USE_XML_WRITER=1
#
# For one off testing we can use the following:
#
#    _EXCEL_WRITER_XLSX_USE_XML_WRITER=1 perl example.pl
#
sub _set_xml_writer {

    my $self     = shift;
    my $filename = shift;

    my $fh = IO::File->new( $filename, 'w' );
    croak "Couldn't open file $filename for writing.\n" unless $fh;

    binmode $fh, ':utf8';

    my $writer;

    if ( $ENV{_EXCEL_WRITER_XLSX_USE_XML_WRITER} ) {
        require XML::Writer;
        $writer = XML::Writer->new( OUTPUT => $fh );
    }
    else {
        $writer = Excel::Writer::XLSX::Package::XMLwriterSimple->new( $fh );

    }

    croak "Couldn't create XML writer object for $filename.\n" unless $writer;

    $self->{_writer} = $writer;
}


###############################################################################
#
# _write_xml_declaration()
#
# Write the XML declaration.
#
sub _write_xml_declaration {

    my $self       = shift;
    my $writer     = $self->{_writer};
    my $encoding   = 'UTF-8';
    my $standalone = 1;

    $writer->xmlDecl( $encoding, $standalone );
}


1;


__END__

=pod

=head1 NAME

XMLwriter - A base class for the Excel::Writer::XLSX writer classes.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Writer::XLSX>.

=cut
