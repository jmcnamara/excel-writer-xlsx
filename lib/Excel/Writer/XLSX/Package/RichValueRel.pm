package Excel::Writer::XLSX::Package::RichValueRel;

###############################################################################
#
# RichValueRel - A class for writing the Excel XLSX richValueRel.xml file.
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

    $self->{_value_count} = 0;

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
    $self->_write_rich_value_rels();

}


##############################################################################
#
# _write_rich_value_rels()
#
# Write the <richValueRels> element.
#
sub _write_rich_value_rels {

    my $self = shift;
    my $xmlns =
      'http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel';
    my $xmlns_r =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    my @attributes = (
        'xmlns'   => $xmlns,
        'xmlns:r' => $xmlns_r,
    );

    $self->xml_start_tag( 'richValueRels', @attributes );


    for my $index ( 0 .. $self->{_value_count} - 1 ) {

        # Write the rel element.
        $self->_write_rel( $index + 1 );
    }


    $self->xml_end_tag( 'richValueRels' );
}


##############################################################################
#
# _write_rel()
#
# Write the <rel> element.
#
sub _write_rel {

    my $self = shift;
    my $r_id = 'rId' . shift;

    my @attributes = ( 'r:id' => $r_id, );

    $self->xml_empty_tag( 'rel', @attributes );
}


1;


__END__

=pod

=head1 NAME

RichValueRel - A class for writing the Excel XLSX richValueRel.xml file.

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
