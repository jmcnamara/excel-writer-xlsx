package Excel::Writer::XLSX::Package::RichValue;

###############################################################################
#
# RichValue - A class for writing the Excel XLSX rdrichValue file.
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

    $self->{_embedded_images} = [];

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

    # Write the rvData element.
    $self->_write_rv_data();

    $self->xml_end_tag( 'rvData' );
}


##############################################################################
#
# _write_rv_data()
#
# Write the <rvData> element.
#
sub _write_rv_data {

    my $self = shift;
    my $xmlns =
      'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';

    my @attributes = (
        'xmlns' => $xmlns,
        'count' => scalar @{ $self->{_embedded_images} },
    );

    $self->xml_start_tag( 'rvData', @attributes );

    my $index = 0;
    for my $image ( @{ $self->{_embedded_images} } ) {

        # Write the rv element.
        $self->_write_rv( $index, $image->[2], $image->[3] );

        $index++;
    }


}


##############################################################################
#
# _write_rv()
#
# Write the <rv> element.
#
sub _write_rv {

    my $self        = shift;
    my $index       = shift;
    my $description = shift;
    my $decorative  = shift;
    my $value       = 5;

    if ( $decorative ) {
        $value = 6;
    }

    my @attributes = ( 's' => 0 );

    $self->xml_start_tag( 'rv', @attributes );

    $self->_write_v( $index );
    $self->_write_v( $value );

    if ( $description ) {
        $self->_write_v( $description );
    }

    $self->xml_end_tag( 'rv' );
}


##############################################################################
#
# _write_v()
#
# Write the <v> element.
#
sub _write_v {

    my $self = shift;
    my $data = shift;

    $self->xml_data_element( 'v', $data );
}


1;


__END__

=pod

=head1 NAME

RichValue - A class for writing the Excel XLSX RichValue file.

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
