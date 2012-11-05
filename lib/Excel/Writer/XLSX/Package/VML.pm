package Excel::Writer::XLSX::Package::VML;

###############################################################################
#
# VML - A class for writing the Excel XLSX VML files.
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
use Carp;
use Excel::Writer::XLSX::Package::XMLwriter;

our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
our $VERSION = '0.54';


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

    my $self          = shift;
    my $data_id       = shift;
    my $vml_shape_id  = shift;
    my $comments_data = shift;

    $self->_write_xml_namespace;

    # Write the o:shapelayout element.
    $self->_write_shapelayout( $data_id );

    # Write the v:shapetype element.
    $self->_write_shapetype();

    my $z_index = 1;
    for my $comment ( @$comments_data ) {

        # Write the v:shape element.
        $self->_write_shape( ++$vml_shape_id, $z_index++, $comment );
    }

    $self->xml_end_tag( 'xml' );

    # Close the XML writer filehandle.
    $self->xml_get_fh()->close();
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _pixels_to_points()
#
# Convert comment vertices from pixels to points.
#
sub _pixels_to_points {

    my $self     = shift;
    my $vertices = shift;

    my (
        $col_start, $row_start, $x1,    $y1,
        $col_end,   $row_end,   $x2,    $y2,
        $left,      $top,       $width, $height
    ) = @$vertices;

    for my $pixels ( $left, $top, $width, $height ) {
        $pixels *= 0.75;
    }

    return ( $left, $top, $width, $height );
}


###############################################################################
#
# XML writing methods.
#
###############################################################################

###############################################################################
#
# _write_xml_namespace()
#
# Write the <xml> element. This is the root element of VML.
#
sub _write_xml_namespace {

    my $self    = shift;
    my $schema  = 'urn:schemas-microsoft-com:';
    my $xmlns   = $schema . 'vml';
    my $xmlns_o = $schema . 'office:office';
    my $xmlns_x = $schema . 'office:excel';

    my @attributes = (
        'xmlns:v' => $xmlns,
        'xmlns:o' => $xmlns_o,
        'xmlns:x' => $xmlns_x,
    );

    $self->xml_start_tag( 'xml', @attributes );
}


##############################################################################
#
# _write_shapelayout()
#
# Write the <o:shapelayout> element.
#
sub _write_shapelayout {

    my $self    = shift;
    my $data_id = shift;
    my $ext     = 'edit';

    my @attributes = ( 'v:ext' => $ext );

    $self->xml_start_tag( 'o:shapelayout', @attributes );

    # Write the o:idmap element.
    $self->_write_idmap( $data_id );

    $self->xml_end_tag( 'o:shapelayout' );
}


##############################################################################
#
# _write_idmap()
#
# Write the <o:idmap> element.
#
sub _write_idmap {

    my $self    = shift;
    my $ext     = 'edit';
    my $data_id = shift;

    my @attributes = (
        'v:ext' => $ext,
        'data'  => $data_id,
    );

    $self->xml_empty_tag( 'o:idmap', @attributes );
}


##############################################################################
#
# _write_shapetype()
#
# Write the <v:shapetype> element.
#
sub _write_shapetype {

    my $self      = shift;
    my $id        = '_x0000_t202';
    my $coordsize = '21600,21600';
    my $spt       = 202;
    my $path      = 'm,l,21600r21600,l21600,xe';

    my @attributes = (
        'id'        => $id,
        'coordsize' => $coordsize,
        'o:spt'     => $spt,
        'path'      => $path,
    );

    $self->xml_start_tag( 'v:shapetype', @attributes );

    # Write the v:stroke element.
    $self->_write_stroke();

    # Write the v:path element.
    $self->_write_path( 't', 'rect' );

    $self->xml_end_tag( 'v:shapetype' );
}


##############################################################################
#
# _write_stroke()
#
# Write the <v:stroke> element.
#
sub _write_stroke {

    my $self      = shift;
    my $joinstyle = 'miter';

    my @attributes = ( 'joinstyle' => $joinstyle );

    $self->xml_empty_tag( 'v:stroke', @attributes );
}


##############################################################################
#
# _write_path()
#
# Write the <v:path> element.
#
sub _write_path {

    my $self            = shift;
    my $gradientshapeok = shift;
    my $connecttype     = shift;
    my @attributes      = ();

    push @attributes, ( 'gradientshapeok' => 't' ) if $gradientshapeok;
    push @attributes, ( 'o:connecttype' => $connecttype );

    $self->xml_empty_tag( 'v:path', @attributes );
}


##############################################################################
#
# _write_shape()
#
# Write the <v:shape> element.
#
sub _write_shape {

    my $self       = shift;
    my $id         = shift;
    my $z_index    = shift;
    my $comment    = shift;
    my $type       = '#_x0000_t202';
    my $insetmode  = 'auto';
    my $visibility = 'hidden';

    # Set the shape index.
    $id = '_x0000_s' . $id;

    # Get the comment parameters
    my $row       = $comment->[0];
    my $col       = $comment->[1];
    my $string    = $comment->[2];
    my $author    = $comment->[3];
    my $visible   = $comment->[4];
    my $fillcolor = $comment->[5];
    my $vertices  = $comment->[6];

    my ( $left, $top, $width, $height ) = $self->_pixels_to_points( $vertices );

    # Set the visibility.
    $visibility = 'visible' if $visible;

    my $style =
        'position:absolute;'
      . 'margin-left:'
      . $left . 'pt;'
      . 'margin-top:'
      . $top . 'pt;'
      . 'width:'
      . $width . 'pt;'
      . 'height:'
      . $height . 'pt;'
      . 'z-index:'
      . $z_index . ';'
      . 'visibility:'
      . $visibility;


    my @attributes = (
        'id'          => $id,
        'type'        => $type,
        'style'       => $style,
        'fillcolor'   => $fillcolor,
        'o:insetmode' => $insetmode,
    );

    $self->xml_start_tag( 'v:shape', @attributes );

    # Write the v:fill element.
    $self->_write_fill();

    # Write the v:shadow element.
    $self->_write_shadow();

    # Write the v:path element.
    $self->_write_path( undef, 'none' );

    # Write the v:textbox element.
    $self->_write_textbox();

    # Write the x:ClientData element.
    $self->_write_client_data( $row, $col, $visible, $vertices );

    $self->xml_end_tag( 'v:shape' );
}


##############################################################################
#
# _write_fill()
#
# Write the <v:fill> element.
#
sub _write_fill {

    my $self    = shift;
    my $color_2 = '#ffffe1';

    my @attributes = ( 'color2' => $color_2 );

    $self->xml_empty_tag( 'v:fill', @attributes );
}


##############################################################################
#
# _write_shadow()
#
# Write the <v:shadow> element.
#
sub _write_shadow {

    my $self     = shift;
    my $on       = 't';
    my $color    = 'black';
    my $obscured = 't';

    my @attributes = (
        'on'       => $on,
        'color'    => $color,
        'obscured' => $obscured,
    );

    $self->xml_empty_tag( 'v:shadow', @attributes );
}


##############################################################################
#
# _write_textbox()
#
# Write the <v:textbox> element.
#
sub _write_textbox {

    my $self  = shift;
    my $style = 'mso-direction-alt:auto';

    my @attributes = ( 'style' => $style );

    $self->xml_start_tag( 'v:textbox', @attributes );

    # Write the div element.
    $self->_write_div();

    $self->xml_end_tag( 'v:textbox' );
}


##############################################################################
#
# _write_div()
#
# Write the <div> element.
#
sub _write_div {

    my $self  = shift;
    my $style = 'text-align:left';

    my @attributes = ( 'style' => $style );

    $self->xml_start_tag( 'div', @attributes );


    $self->xml_end_tag( 'div' );
}


##############################################################################
#
# _write_client_data()
#
# Write the <x:ClientData> element.
#
sub _write_client_data {

    my $self        = shift;
    my $row         = shift;
    my $col         = shift;
    my $visible     = shift;
    my $vertices    = shift;
    my $object_type = 'Note';

    my @attributes = ( 'ObjectType' => $object_type );

    $self->xml_start_tag( 'x:ClientData', @attributes );

    # Write the x:MoveWithCells element.
    $self->_write_move_with_cells();

    # Write the x:SizeWithCells element.
    $self->_write_size_with_cells();

    # Write the x:Anchor element.
    $self->_write_anchor( $vertices );

    # Write the x:AutoFill element.
    $self->_write_auto_fill();

    # Write the x:Row element.
    $self->_write_row( $row );

    # Write the x:Column element.
    $self->_write_column( $col );

    # Write the x:Visible element.
    $self->_write_visible() if $visible;

    $self->xml_end_tag( 'x:ClientData' );
}


##############################################################################
#
# _write_move_with_cells()
#
# Write the <x:MoveWithCells> element.
#
sub _write_move_with_cells {

    my $self = shift;

    $self->xml_empty_tag( 'x:MoveWithCells' );
}


##############################################################################
#
# _write_size_with_cells()
#
# Write the <x:SizeWithCells> element.
#
sub _write_size_with_cells {

    my $self = shift;

    $self->xml_empty_tag( 'x:SizeWithCells' );
}


##############################################################################
#
# _write_visible()
#
# Write the <x:Visible> element.
#
sub _write_visible {

    my $self = shift;

    $self->xml_empty_tag( 'x:Visible' );
}


##############################################################################
#
# _write_anchor()
#
# Write the <x:Anchor> element.
#
sub _write_anchor {

    my $self     = shift;
    my $vertices = shift;

    my ( $col_start, $row_start, $x1, $y1, $col_end, $row_end, $x2, $y2 ) =
      @$vertices;

    my $data = join ", ",
      ( $col_start, $x1, $row_start, $y1, $col_end, $x2, $row_end, $y2 );

    $self->xml_data_element( 'x:Anchor', $data );
}


##############################################################################
#
# _write_auto_fill()
#
# Write the <x:AutoFill> element.
#
sub _write_auto_fill {

    my $self = shift;
    my $data = 'False';

    $self->xml_data_element( 'x:AutoFill', $data );
}


##############################################################################
#
# _write_row()
#
# Write the <x:Row> element.
#
sub _write_row {

    my $self = shift;
    my $data = shift;

    $self->xml_data_element( 'x:Row', $data );
}


##############################################################################
#
# _write_column()
#
# Write the <x:Column> element.
#
sub _write_column {

    my $self = shift;
    my $data = shift;

    $self->xml_data_element( 'x:Column', $data );
}


1;


__END__

=pod

=head1 NAME

VML - A class for writing the Excel XLSX VML files.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

(c) MM-MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Writer::XLSX>.

=cut
