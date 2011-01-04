package Excel::Writer::XLSX::Package::Core;

###############################################################################
#
# Core - A class for writing the Excel XLSX core.xml file.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2010, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.010000;
use strict;
use warnings;
use Carp;
use Excel::Writer::XLSX::Package::XMLwriter;

our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
our $VERSION = '0.05';


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
    my $self  = Excel::Writer::XLSX::Package::XMLwriter->new();

    $self->{_writer}            = undef;
    $self->{_creator}           = '';
    $self->{_modifier}          = '';
    $self->{_creation_date}     = '2010-01-01T00:00:00Z';
    $self->{_modification_date} = '2010-01-01T00:00:00Z';

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
    $self->_write_cp_core_properties();
    $self->_write_dc_creator();
    $self->_write_cp_last_modified_by();
    $self->_write_dcterms_created();
    $self->_write_dcterms_modified();

    $self->{_writer}->endTag( 'cp:coreProperties' );

    # Close the XM writer object and filehandle.
    $self->{_writer}->end();
    $self->{_writer}->getOutput()->close();
}


###############################################################################
#
# _set_creator()
#
# Set the document creator.
#
sub _set_creator {
    my $self    = shift;
    my $creator = shift;

    $self->{_creator} = $creator;
}


###############################################################################
#
# _set_modifier()
#
# Set the document modifier.
#
sub _set_modifier {
    my $self     = shift;
    my $modifier = shift;

    $self->{_modifier} = $modifier;
}


###############################################################################
#
# _set_creation_date()
#
# Set the document creation date.
#
sub _set_creation_date {
    my $self          = shift;
    my $creation_date = shift;

    $self->{_creation_date} = $creation_date;
}


###############################################################################
#
# _set_modification_date()
#
# Set the document modification date.
#
sub _set_modification_date {
    my $self              = shift;
    my $modification_date = shift;

    $self->{_modification_date} = $modification_date;
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_cp_core_properties()
#
# Write the <cp:coreProperties> element.
#
sub _write_cp_core_properties {

    my $self = shift;
    my $xmlns_cp =
      'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
    my $xmlns_dc       = 'http://purl.org/dc/elements/1.1/';
    my $xmlns_dcterms  = 'http://purl.org/dc/terms/';
    my $xmlns_dcmitype = 'http://purl.org/dc/dcmitype/';
    my $xmlns_xsi      = 'http://www.w3.org/2001/XMLSchema-instance';

    my @attributes = (
        'xmlns:cp'       => $xmlns_cp,
        'xmlns:dc'       => $xmlns_dc,
        'xmlns:dcterms'  => $xmlns_dcterms,
        'xmlns:dcmitype' => $xmlns_dcmitype,
        'xmlns:xsi'      => $xmlns_xsi,
    );

    $self->{_writer}->startTag( 'cp:coreProperties', @attributes );
}


###############################################################################
#
# _write_dc_creator()
#
# Write the <dc:creator> element.
#
sub _write_dc_creator {

    my $self = shift;
    my $data = $self->{_creator};

    $self->{_writer}->dataElement( 'dc:creator', $data );
}


###############################################################################
#
# _write_cp_last_modified_by()
#
# Write the <cp:lastModifiedBy> element.
#
sub _write_cp_last_modified_by {

    my $self = shift;
    my $data = $self->{_modifier};

    $self->{_writer}->dataElement( 'cp:lastModifiedBy', $data );
}


###############################################################################
#
# _write_dcterms_created()
#
# Write the <dcterms:created> element.
#
sub _write_dcterms_created {

    my $self     = shift;
    my $data     = $self->{_creation_date};
    my $xsi_type = 'dcterms:W3CDTF';

    my @attributes = ( 'xsi:type' => $xsi_type, );


    $self->{_writer}->dataElement( 'dcterms:created', $data, @attributes );
}


###############################################################################
#
# _write_dcterms_modified()
#
# Write the <dcterms:modified> element.
#
sub _write_dcterms_modified {

    my $self     = shift;
    my $data     = $self->{_modification_date};
    my $xsi_type = 'dcterms:W3CDTF';

    my @attributes = ( 'xsi:type' => $xsi_type, );


    $self->{_writer}->dataElement( 'dcterms:modified', $data, @attributes );
}


1;


__END__

=pod

=head1 NAME

Core - A class for writing the Excel XLSX core.xml file.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

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
