package Excel::Writer::XLSX::Package::Metadata;

###############################################################################
#
# Metadata - A class for writing the Excel XLSX metadata.xml file.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2024, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.008002;
use strict;
use warnings;
use Carp;
use Excel::Writer::XLSX::Package::XMLwriter;

our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
our $VERSION = '1.14';


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
    my $fh    = shift;
    my $self  = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );

    $self->{_has_dynamic_functions} = 0;
    $self->{_has_embedded_images}   = 0;
    $self->{_num_embedded_images}   = 0;

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

    if ( $self->{_num_embedded_images} > 0 ) {
        $self->{_has_embedded_images} = 1;
    }

    $self->xml_declaration;

    # Write the metadata element.
    $self->_write_metadata();

    # Write the metadataTypes element.
    $self->_write_metadata_types();

    # Write the futureMetadata element.
    $self->_write_cell_future_metadata()  if $self->{_has_dynamic_functions};
    $self->_write_value_future_metadata() if $self->{_has_embedded_images};

    # Write the valueMetadata element.
    $self->_write_cell_metadata()  if $self->{_has_dynamic_functions};
    $self->_write_value_metadata() if $self->{_has_embedded_images};

    $self->xml_end_tag( 'metadata' );

    # Close the XML writer filehandle.
    $self->xml_get_fh()->close();
}


###############################################################################
#
# XML writing methods.
#
###############################################################################


##############################################################################
#
# _write_metadata()
#
# Write the <metadata> element.
#
sub _write_metadata {

    my $self = shift;

    my @attributes = (
        'xmlns' =>
        'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    );


    if ( $self->{_has_embedded_images} ) {
        push @attributes,
          (
              'xmlns:xlrd' =>
              'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata'
          );
    }

    if ( $self->{_has_dynamic_functions} ) {
        push @attributes,
          (
              'xmlns:xda' =>
              'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray'
          );
    }

    $self->xml_start_tag( 'metadata', @attributes );
}


##############################################################################
#
# _write_metadata_types()
#
# Write the <metadataTypes> element.
#
sub _write_metadata_types {

    my $self  = shift;
    my $count = $self->{_has_dynamic_functions} + $self->{_has_embedded_images};

    my @attributes = ( 'count' => $count );

    $self->xml_start_tag( 'metadataTypes', @attributes );

    # Write the metadataType element.
    $self->_write_cell_metadata_type()  if $self->{_has_dynamic_functions};
    $self->_write_value_metadata_type() if $self->{_has_embedded_images};

    $self->xml_end_tag( 'metadataTypes' );
}


##############################################################################
#
# _write_cell_metadata_type()
#
# Write the <metadataType> element.
#
sub _write_cell_metadata_type {

    my $self = shift;

    my @attributes = (
        'name'                => 'XLDAPR',
        'minSupportedVersion' => 120000,
        'copy'                => 1,
        'pasteAll'            => 1,
        'pasteValues'         => 1,
        'merge'               => 1,
        'splitFirst'          => 1,
        'rowColShift'         => 1,
        'clearFormats'        => 1,
        'clearComments'       => 1,
        'assign'              => 1,
        'coerce'              => 1,
        'cellMeta'            => 1,
    );

    $self->xml_empty_tag( 'metadataType', @attributes );
}


##############################################################################
#
# _write_value_metadata_type()
#
# Write the <metadataType> element.
#
sub _write_value_metadata_type {

    my $self = shift;

    my @attributes = (
        'name'                => 'XLRICHVALUE',
        'minSupportedVersion' => 120000,
        'copy'                => 1,
        'pasteAll'            => 1,
        'pasteValues'         => 1,
        'merge'               => 1,
        'splitFirst'          => 1,
        'rowColShift'         => 1,
        'clearFormats'        => 1,
        'clearComments'       => 1,
        'assign'              => 1,
        'coerce'              => 1,
    );

    $self->xml_empty_tag( 'metadataType', @attributes );
}


##############################################################################
#
# _write_cell_future_metadata()
#
# Write the <futureMetadata> element.
#
sub _write_cell_future_metadata {

    my $self  = shift;

    my @attributes = (
        'name'  => 'XLDAPR',
        'count' => 1,
    );

    $self->xml_start_tag( 'futureMetadata', @attributes );
    $self->xml_start_tag( 'bk' );
    $self->xml_start_tag( 'extLst' );

    # Write the ext element.
    $self->_write_cell_ext();

    $self->xml_end_tag( 'ext' );
    $self->xml_end_tag( 'extLst' );
    $self->xml_end_tag( 'bk' );

    $self->xml_end_tag( 'futureMetadata' );
}


##############################################################################
#
# _write_value_future_metadata()
#
# Write the <futureMetadata> element.
#
sub _write_value_future_metadata {

    my $self       = shift;
    my $num_images = $self->{_num_embedded_images};

    my @attributes = (
        'name'  => 'XLRICHVALUE',
        'count' => $num_images,
    );

    $self->xml_start_tag( 'futureMetadata', @attributes );

    for my $i ( 0 .. $num_images - 1 ) {
        $self->xml_start_tag( 'bk' );
        $self->xml_start_tag( 'extLst' );

        # Write the ext element.
        $self->_write_value_ext( $i );

        $self->xml_end_tag( 'ext' );
        $self->xml_end_tag( 'extLst' );
        $self->xml_end_tag( 'bk' );
    }

    $self->xml_end_tag( 'futureMetadata' );
}


##############################################################################
#
# _write_cell_ext()
#
# Write the <ext> element.
#
sub _write_cell_ext {

    my $self = shift;
    my $uri = shift;

    my @attributes = ( 'uri' => '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}' );

    $self->xml_start_tag( 'ext', @attributes );

    # Write the xda:dynamicArrayProperties element.
    $self->_write_xda_dynamic_array_properties();
}


##############################################################################
#
# _write_value_ext()
#
# Write the <ext> element.
#
sub _write_value_ext {

    my $self = shift;
    my $num  = shift;

    my @attributes = ( 'uri' => '{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}' );

    $self->xml_start_tag( 'ext', @attributes );

    # Write the <xlrd:rvb> element.
    $self->_write_xlrd_rvb( $num );
}

##############################################################################
#
# _write_xda_dynamic_array_properties()
#
# Write the <xda:dynamicArrayProperties> element.
#
sub _write_xda_dynamic_array_properties {

    my $self        = shift;

    my @attributes = (
        'fDynamic'   => 1,
        'fCollapsed' => 0,
    );

    $self->xml_empty_tag( 'xda:dynamicArrayProperties', @attributes );
}


##############################################################################
#
# _write_cell_metadata()
#
# Write the <cellMetadata> element.
#
sub _write_cell_metadata {

    my $self  = shift;
    my $count = 1;

    my @attributes = ( 'count' => $count, );

    $self->xml_start_tag( 'cellMetadata', @attributes );
    $self->xml_start_tag( 'bk' );

    # Write the rc element.
    $self->_write_rc(1, 0);

    $self->xml_end_tag( 'bk' );
    $self->xml_end_tag( 'cellMetadata' );
}


##############################################################################
#
# _write_value_metadata()
#
# Write the <valueMetadata> element.
#
sub _write_value_metadata {

    my $self  = shift;
    my $count = $self->{_num_embedded_images};
    my $type  = 1;

    if ($self->{_has_dynamic_functions}) {
        $type  = 2;
    }
    my @attributes = ( 'count' => $count, );

    $self->xml_start_tag( 'valueMetadata', @attributes );

    for my $i ( 0 .. $count - 1 ) {
        $self->xml_start_tag( 'bk' );

        # Write the rc element.
        $self->_write_rc( $type, $i );

        $self->xml_end_tag( 'bk' );
    }

    $self->xml_end_tag( 'valueMetadata' );
}


##############################################################################
#
# _write_rc()
#
# Write the <rc> element.
#
sub _write_rc {

    my $self  = shift;
    my $type  = shift;
    my $value = shift;

    my @attributes = (
        't' => $type,
        'v' => $value,
    );

    $self->xml_empty_tag( 'rc', @attributes );
}


##############################################################################
#
# _write_xlrd_rvb()
#
# Write the <xlrd:rvb> element.
#
sub _write_xlrd_rvb {

    my $self  = shift;
    my $value = shift;

    my @attributes = ( 'i' => $value, );

    $self->xml_empty_tag( 'xlrd:rvb', @attributes );
}


1;


__END__

=pod

=head1 NAME

Metadata - A class for writing the Excel XLSX metadata.xml file.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT


(c) MM-MMXXIV, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<https://dev.perl.org/licenses/artistic.html> or the GNU General Public License v1.0 or later L<https://dev.perl.org/licenses/gpl1.html>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Writer::XLSX>.

=cut
