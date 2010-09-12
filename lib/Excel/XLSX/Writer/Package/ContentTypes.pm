package Excel::XLSX::Writer::Package::ContentTypes;

###############################################################################
#
# Excel::XLSX::Writer::Package::ContentTypes - A class for writing the Excel
# XLS [Content_Types] file.
#
# Used in conjunction with Excel::XLSX::Writer
#
# Copyright 2000-2010, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.010000;
use strict;
use warnings;
use Exporter;
use Carp;
use XML::Writer;
use IO::File;

our @ISA     = qw(Exporter);
our $VERSION = '0.01';


###############################################################################
#
# Package data.
#
###############################################################################


our @defaults = (
    [ 'xml',  'application/xml' ],
    [ 'jpeg', 'image/jpeg' ],
    [ 'rels', 'application/vnd.openxmlformats-package.relationships+xml' ],
);

our @overrides = (
    #<<<
    [
        '/docProps/app.xml',
        'application/vnd.openxmlformats-officedocument.extended-properties+xml'
    ],
    [
        '/docProps/core.xml',
        'application/vnd.openxmlformats-package.core-properties+xml'
    ],
    [
        '/xl/workbook.xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
    ],
    [
        '/xl/styles.xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
    ],
    [
        '/xl/theme/theme1.xml',
        'application/vnd.openxmlformats-officedocument.theme+xml'
    ],
    #>>>
);


###############################################################################
#
# Public and private API methods.
#
###############################################################################

###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;

    my $self = {
        _writer    => undef,
        _defaults  => \@defaults,
        _overrides => \@overrides,
    };

    bless $self, $class;

    return $self;
}


###############################################################################
#
# _assemble_xml_file()
#
# Assemble and write the XML file.
#
sub _assemble_xml_file {

    my $self = shift;

    return unless $self->{_writer};

    $self->_write_xml_declaration;
    $self->_write_types();
    $self->_write_defaults();
    $self->_write_overrides();

    $self->{_writer}->endTag( 'Types' );
}


###############################################################################
#
# _set_xml_writer()
#
# Set the XML::Writer for the object.
#
sub _set_xml_writer {

    my $self     = shift;
    my $filename = shift;

    my $output = new IO::File( $filename, 'w' );
    croak "Couldn't open file $filename for writing.\n" unless $output;

    my $writer = new XML::Writer( OUTPUT => $output );
    croak "Couldn't create XML::Writer for $filename.\n" unless $writer;

    $self->{_writer} = $writer;
}


###############################################################################
#
# _add_override()
#
# Add elements to the ContentTypes overrides.
#
sub _add_override {

    my $self         = shift;
    my $part_name    = shift;
    my $content_type = shift;

    push @{ $self->{_overrides} }, [ $part_name, $content_type ];

}


###############################################################################
#
# _add_sheet_name()
#
# Add the name of a worksheet to the ContentTypes overrides.
#
sub _add_sheet_name {

    my $self       = shift;
    my $sheet_name = shift;

    $sheet_name = "/xl/worksheets/$sheet_name.xml";

    $self->_add_override( $sheet_name,
            'application/vnd.openxmlformats-officedocument.'
          . 'spreadsheetml.worksheet+xml' );
}


###############################################################################
#
# _Add_shared_strings()
#
# Add the sharedStrings link to the ContentTypes overrides.
#
sub _add_shared_strings {

    my $self = shift;

    $self->_add_override( '/xl/sharedStrings.xml',
            'application/vnd.openxmlformats-officedocument.'
          . 'spreadsheetml.sharedStrings+xml' );
}


###############################################################################
#
# _add_calc_chain()
#
# Add the calcChain link to the ContentTypes overrides.
#
sub _add_calc_chain {

    my $self = shift;

    $self->_add_override( '/xl/calcChain.xml',
            'application/vnd.openxmlformats-officedocument.'
          . 'spreadsheetml.calcChain+xml' );
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _write_defaults()
#
# Write out all of the <Default> types.
#
sub _write_defaults {

    my $self = shift;

    for my $aref ( @{ $self->{_defaults} } ) {
        #<<<
        $self->{_writer}->emptyTag(
            'Default',
            'Extension',   $aref->[0],
            'ContentType', $aref->[1] );
        #>>>
    }
}


###############################################################################
#
# _write_overrides()
#
# Write out all of the <Override> types.
#
sub _write_overrides {

    my $self = shift;

    for my $aref ( @{ $self->{_overrides} } ) {
        #<<<
        $self->{_writer}->emptyTag(
            'Override',
            'PartName',    $aref->[0],
            'ContentType', $aref->[1] );
        #>>>
    }
}


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_xml_declaration()
#
# Write the XML declaration.
#
sub _write_xml_declaration {

    my $self       = shift;
    my $encoding   = 'UTF-8';
    my $standalone = 1;

    $self->{_writer}->xmlDecl( $encoding, $standalone );
}

###############################################################################
#
# _write_types()
#
# Write the <Types> element.
#
sub _write_types {

    my $self  = shift;
    my $xmlns = 'http://schemas.openxmlformats.org/package/2006/content-types';

    my @attributes = ( 'xmlns' => $xmlns, );

    $self->{_writer}->startTag( 'Types', @attributes );
}

###############################################################################
#
# _write_default()
#
# Write the <Default> element.
#
sub _write_default {

    my $self         = shift;
    my $extension    = shift;
    my $content_type = shift;

    my @attributes = (
        'Extension'   => $extension,
        'ContentType' => $content_type,
    );

    $self->{_writer}->emptyTag( 'Default', @attributes );
}


###############################################################################
#
# _write_override()
#
# Write the <Override> element.
#
sub _write_override {

    my $self         = shift;
    my $part_name    = shift;
    my $content_type = shift;
    my $writer       = $self->{_writer};

    my @attributes = (
        'PartName'    => $part_name,
        'ContentType' => $content_type,
    );

    $self->{_writer}->emptyTag( 'Override', @attributes );
}


1;


__END__

=pod

=head1 NAME

Excel::XLSX::Writer::Package::ContentTypes - A class for writing the Excel XLSX [Content_Types] file.

=head1 SYNOPSIS

See the documentation for L<Excel::XLSX::Writer>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::XLSX::Writer>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::XLSX::Writer>.

=cut
