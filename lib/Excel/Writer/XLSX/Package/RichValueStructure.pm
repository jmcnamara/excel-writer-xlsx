package Excel::Writer::XLSX::Package::RichValueStructure;

###############################################################################
#
# RichValueStructure - A class for writing the Excel XLSX rdrichvaluestructure.xml
# file.
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
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $fh    = shift;
    my $self  = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );

    $self->{_has_embedded_descriptions} = 0;

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

    $self->xml_declaration;
    $self->_write_rv_structures();

    $self->xml_end_tag( 'rvStructures' );

}


##############################################################################
#
# _write_rv_structures()
#
# Write the <rvStructures> element.
#
sub _write_rv_structures {

    my $self = shift;
    my $xmlns =
      'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
    my $count = 1;

    my @attributes = (
        'xmlns' => $xmlns,
        'count' => $count,
    );

    $self->xml_start_tag( 'rvStructures', @attributes );


    # Write the s element.
    $self->_write_s();

}


##############################################################################
#
# _write_s()
#
# Write the <s> element.
#
sub _write_s {

    my $self = shift;
    my $t    = '_localImage';

    my @attributes = ( 't' => $t, );

    $self->xml_start_tag( 's', @attributes );

    $self->_write_k( '_rvRel:LocalImageIdentifier', 'i' );
    $self->_write_k( 'CalcOrigin',                  'i' );

    if ( $self->{_has_embedded_descriptions} ) {
        $self->_write_k( 'Text', 's' );
    }

    $self->xml_end_tag( 's' );
}


##############################################################################
#
# _write_k()
#
# Write the <k> element.
#
sub _write_k {

    my $self = shift;
    my $n    = shift;
    my $t    = shift;

    my @attributes = (
        'n' => $n,
        't' => $t,
    );

    $self->xml_empty_tag( 'k', @attributes );
}


1;


__END__

=pod

=head1 NAME

RichValueStructure - A class for writing the Excel XLSX rdrichvaluestructure.xml file.

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
