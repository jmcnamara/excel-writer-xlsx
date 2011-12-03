package Excel::Writer::XLSX::Drawing;

###############################################################################
#
# Drawing - A class for writing the Excel XLSX drawing.xml file.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2011, John McNamara, jmcnamara@cpan.org
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
our $VERSION = '0.39';


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

    $self->{_writer}      = undef;
    $self->{_drawings}    = [];
    $self->{_embedded}    = 0;
    $self->{_orientation} = 0;

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

    # Write the xdr:wsDr element.
    $self->_write_drawing_workspace();

    if ( $self->{_embedded} ) {

        my $index = 0;
        for my $dimensions ( @{ $self->{_drawings} } ) {

            # Write the xdr:twoCellAnchor element.
            $self->_write_two_cell_anchor( ++$index, @$dimensions );
        }

    }
    else {
        my $index = 0;

        # Write the xdr:absoluteAnchor element.
        $self->_write_absolute_anchor( ++$index );
    }

    $self->{_writer}->endTag( 'xdr:wsDr' );

    # Close the XM writer object and filehandle.
    $self->{_writer}->end();
    $self->{_writer}->getOutput()->close();
}


###############################################################################
#
# _add_drawing_object()
#
# Add a chart or image sub object to the drawing.
#
sub _add_drawing_object {

    my $self = shift;

    push @{ $self->{_drawings} }, [@_];
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


##############################################################################
#
# _write_drawing_workspace()
#
# Write the <xdr:wsDr> element.
#
sub _write_drawing_workspace {

    my $self      = shift;
    my $schema    = 'http://schemas.openxmlformats.org/drawingml/';
    my $xmlns_xdr = $schema . '2006/spreadsheetDrawing';
    my $xmlns_a   = $schema . '2006/main';

    my @attributes = (
        'xmlns:xdr' => $xmlns_xdr,
        'xmlns:a'   => $xmlns_a,
    );

    $self->{_writer}->startTag( 'xdr:wsDr', @attributes );
}


##############################################################################
#
# _write_two_cell_anchor()
#
# Write the <xdr:twoCellAnchor> element.
#
sub _write_two_cell_anchor {

    my $self            = shift;
    my $index           = shift;
    my $type            = shift;
    my $col_from        = shift;
    my $row_from        = shift;
    my $col_from_offset = shift;
    my $row_from_offset = shift;
    my $col_to          = shift;
    my $row_to          = shift;
    my $col_to_offset   = shift;
    my $row_to_offset   = shift;
    my $col_absolute    = shift;
    my $row_absolute    = shift;
    my $width           = shift;
    my $height          = shift;
    my $description     = shift;
    my @attributes      = ();


    # Add attribute for images.
    if ( $type == 2 ) {
        push @attributes, ( editAs => 'oneCell' );
    }

    $self->{_writer}->startTag( 'xdr:twoCellAnchor', @attributes );

    # Write the xdr:from element.
    $self->_write_from(
        $col_from,
        $row_from,
        $col_from_offset,
        $row_from_offset,

    );

    # Write the xdr:from element.
    $self->_write_to(
        $col_to,
        $row_to,
        $col_to_offset,
        $row_to_offset,

    );

    if ( $type == 1 ) {

        # Write the xdr:graphicFrame element for charts.
        $self->_write_graphic_frame( $index );
    }
    else {

        # Write the xdr:pic element.
        $self->_write_pic( $index, $col_absolute, $row_absolute, $width,
            $height, $description );
    }

    # Write the xdr:clientData element.
    $self->_write_client_data();

    $self->{_writer}->endTag( 'xdr:twoCellAnchor' );
}


##############################################################################
#
# _write_absolute_anchor()
#
# Write the <xdr:absoluteAnchor> element.
#
sub _write_absolute_anchor {

    my $self  = shift;
    my $index = shift;

    $self->{_writer}->startTag( 'xdr:absoluteAnchor' );

    # Different co-ordinates for horizonatal (= 0) and vertical (= 1).
    if ( $self->{_orientation} == 0 ) {

        # Write the xdr:pos element.
        $self->_write_pos( 0, 0 );

        # Write the xdr:ext element.
        $self->_write_ext( 9308969, 6078325 );

    }
    else {

        # Write the xdr:pos element.
        $self->_write_pos( 0, -47625 );

        # Write the xdr:ext element.
        $self->_write_ext( 6162675, 6124575 );

    }


    # Write the xdr:graphicFrame element.
    $self->_write_graphic_frame( $index );

    # Write the xdr:clientData element.
    $self->_write_client_data();

    $self->{_writer}->endTag( 'xdr:absoluteAnchor' );
}


##############################################################################
#
# _write_from()
#
# Write the <xdr:from> element.
#
sub _write_from {

    my $self       = shift;
    my $col        = shift;
    my $row        = shift;
    my $col_offset = shift;
    my $row_offset = shift;

    $self->{_writer}->startTag( 'xdr:from' );

    # Write the xdr:col element.
    $self->_write_col( $col );

    # Write the xdr:colOff element.
    $self->_write_col_off( $col_offset );

    # Write the xdr:row element.
    $self->_write_row( $row );

    # Write the xdr:rowOff element.
    $self->_write_row_off( $row_offset );

    $self->{_writer}->endTag( 'xdr:from' );
}


##############################################################################
#
# _write_to()
#
# Write the <xdr:to> element.
#
sub _write_to {

    my $self       = shift;
    my $col        = shift;
    my $row        = shift;
    my $col_offset = shift;
    my $row_offset = shift;

    $self->{_writer}->startTag( 'xdr:to' );

    # Write the xdr:col element.
    $self->_write_col( $col );

    # Write the xdr:colOff element.
    $self->_write_col_off( $col_offset );

    # Write the xdr:row element.
    $self->_write_row( $row );

    # Write the xdr:rowOff element.
    $self->_write_row_off( $row_offset );

    $self->{_writer}->endTag( 'xdr:to' );
}


##############################################################################
#
# _write_col()
#
# Write the <xdr:col> element.
#
sub _write_col {

    my $self = shift;
    my $data = shift;

    $self->{_writer}->dataElement( 'xdr:col', $data );
}


##############################################################################
#
# _write_col_off()
#
# Write the <xdr:colOff> element.
#
sub _write_col_off {

    my $self = shift;
    my $data = shift;

    $self->{_writer}->dataElement( 'xdr:colOff', $data );
}


##############################################################################
#
# _write_row()
#
# Write the <xdr:row> element.
#
sub _write_row {

    my $self = shift;
    my $data = shift;

    $self->{_writer}->dataElement( 'xdr:row', $data );
}


##############################################################################
#
# _write_row_off()
#
# Write the <xdr:rowOff> element.
#
sub _write_row_off {

    my $self = shift;
    my $data = shift;

    $self->{_writer}->dataElement( 'xdr:rowOff', $data );
}


##############################################################################
#
# _write_pos()
#
# Write the <xdr:pos> element.
#
sub _write_pos {

    my $self = shift;
    my $x    = shift;
    my $y    = shift;

    my @attributes = (
        'x' => $x,
        'y' => $y,
    );

    $self->{_writer}->emptyTag( 'xdr:pos', @attributes );
}


##############################################################################
#
# _write_ext()
#
# Write the <xdr:ext> element.
#
sub _write_ext {

    my $self = shift;
    my $cx   = shift;
    my $cy   = shift;

    my @attributes = (
        'cx' => $cx,
        'cy' => $cy,
    );

    $self->{_writer}->emptyTag( 'xdr:ext', @attributes );
}


##############################################################################
#
# _write_graphic_frame()
#
# Write the <xdr:graphicFrame> element.
#
sub _write_graphic_frame {

    my $self   = shift;
    my $index  = shift;
    my $macro  = '';

    my @attributes = ( 'macro' => $macro );

    $self->{_writer}->startTag( 'xdr:graphicFrame', @attributes );

    # Write the xdr:nvGraphicFramePr element.
    $self->_write_nv_graphic_frame_pr( $index );

    # Write the xdr:xfrm element.
    $self->_write_xfrm();

    # Write the a:graphic element.
    $self->_write_atag_graphic( $index );

    $self->{_writer}->endTag( 'xdr:graphicFrame' );
}


##############################################################################
#
# _write_nv_graphic_frame_pr()
#
# Write the <xdr:nvGraphicFramePr> element.
#
sub _write_nv_graphic_frame_pr {

    my $self  = shift;
    my $index  = shift;

    $self->{_writer}->startTag( 'xdr:nvGraphicFramePr' );

    # Write the xdr:cNvPr element.
    $self->_write_c_nv_pr( $index + 1, 'Chart ' . $index );

    # Write the xdr:cNvGraphicFramePr element.
    $self->_write_c_nv_graphic_frame_pr();

    $self->{_writer}->endTag( 'xdr:nvGraphicFramePr' );
}


##############################################################################
#
# _write_c_nv_pr()
#
# Write the <xdr:cNvPr> element.
#
sub _write_c_nv_pr {

    my $self  = shift;
    my $id    = shift;
    my $name  = shift;
    my $descr = shift;

    my @attributes = (
        'id'   => $id,
        'name' => $name,
    );

    # Add description attribute for images.
    if ( defined $descr ) {
        push @attributes, ( descr => $descr );
    }

    $self->{_writer}->emptyTag( 'xdr:cNvPr', @attributes );
}


##############################################################################
#
# _write_c_nv_graphic_frame_pr()
#
# Write the <xdr:cNvGraphicFramePr> element.
#
sub _write_c_nv_graphic_frame_pr {

    my $self = shift;

    if ( $self->{_embedded} ) {
        $self->{_writer}->emptyTag( 'xdr:cNvGraphicFramePr' );
    }
    else {
        $self->{_writer}->startTag( 'xdr:cNvGraphicFramePr' );

        # Write the a:graphicFrameLocks element.
        $self->_write_a_graphic_frame_locks();

        $self->{_writer}->endTag( 'xdr:cNvGraphicFramePr' );
    }
}


##############################################################################
#
# _write_a_graphic_frame_locks()
#
# Write the <a:graphicFrameLocks> element.
#
sub _write_a_graphic_frame_locks {

    my $self   = shift;
    my $no_grp = 1;

    my @attributes = ( 'noGrp' => $no_grp );

    $self->{_writer}->emptyTag( 'a:graphicFrameLocks', @attributes );
}


##############################################################################
#
# _write_xfrm()
#
# Write the <xdr:xfrm> element.
#
sub _write_xfrm {

    my $self = shift;

    $self->{_writer}->startTag( 'xdr:xfrm' );

    # Write the xfrmOffset element.
    $self->_write_xfrm_offset();

    # Write the xfrmOffset element.
    $self->_write_xfrm_extension();

    $self->{_writer}->endTag( 'xdr:xfrm' );
}


##############################################################################
#
# _write_xfrm_offset()
#
# Write the <a:off> xfrm sub-element.
#
sub _write_xfrm_offset {

    my $self = shift;
    my $x    = 0;
    my $y    = 0;

    my @attributes = (
        'x' => $x,
        'y' => $y,
    );

    $self->{_writer}->emptyTag( 'a:off', @attributes );
}


##############################################################################
#
# _write_xfrm_extension()
#
# Write the <a:ext> xfrm sub-element.
#
sub _write_xfrm_extension {

    my $self = shift;
    my $x    = 0;
    my $y    = 0;

    my @attributes = (
        'cx' => $x,
        'cy' => $y,
    );

    $self->{_writer}->emptyTag( 'a:ext', @attributes );
}


##############################################################################
#
# _write_atag_graphic()
#
# Write the <a:graphic> element.
#
sub _write_atag_graphic {

    my $self  = shift;
    my $index = shift;

    $self->{_writer}->startTag( 'a:graphic' );

    # Write the a:graphicData element.
    $self->_write_atag_graphic_data( $index );

    $self->{_writer}->endTag( 'a:graphic' );
}


##############################################################################
#
# _write_atag_graphic_data()
#
# Write the <a:graphicData> element.
#
sub _write_atag_graphic_data {

    my $self  = shift;
    my $index = shift;
    my $uri   = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

    my @attributes = ( 'uri' => $uri, );

    $self->{_writer}->startTag( 'a:graphicData', @attributes );

    # Write the c:chart element.
    $self->_write_c_chart( 'rId' . $index );

    $self->{_writer}->endTag( 'a:graphicData' );
}


##############################################################################
#
# _write_c_chart()
#
# Write the <c:chart> element.
#
sub _write_c_chart {

    my $self    = shift;
    my $r_id    = shift;
    my $schema  = 'http://schemas.openxmlformats.org/';
    my $xmlns_c = $schema . 'drawingml/2006/chart';
    my $xmlns_r = $schema . 'officeDocument/2006/relationships';


    my @attributes = (
        'xmlns:c' => $xmlns_c,
        'xmlns:r' => $xmlns_r,
        'r:id'    => $r_id,
    );

    $self->{_writer}->emptyTag( 'c:chart', @attributes );
}


##############################################################################
#
# _write_client_data()
#
# Write the <xdr:clientData> element.
#
sub _write_client_data {

    my $self = shift;

    $self->{_writer}->emptyTag( 'xdr:clientData' );
}


##############################################################################
#
# _write_pic()
#
# Write the <xdr:pic> element.
#
sub _write_pic {

    my $self         = shift;
    my $index        = shift;
    my $col_absolute = shift;
    my $row_absolute = shift;
    my $width        = shift;
    my $height       = shift;
    my $description  = shift;

    $self->{_writer}->startTag( 'xdr:pic' );

    # Write the xdr:nvPicPr element.
    $self->_write_nv_pic_pr( $index, $description );

    # Write the xdr:blipFill element.
    $self->_write_blip_fill( $index );

    # Write the xdr:spPr element.
    $self->_write_sp_pr( $col_absolute, $row_absolute, $width, $height );

    $self->{_writer}->endTag( 'xdr:pic' );
}


##############################################################################
#
# _write_nv_pic_pr()
#
# Write the <xdr:nvPicPr> element.
#
sub _write_nv_pic_pr {

    my $self        = shift;
    my $index       = shift;
    my $description = shift;

    $self->{_writer}->startTag( 'xdr:nvPicPr' );

    # Write the xdr:cNvPr element.
    $self->_write_c_nv_pr( $index + 1, 'Picture ' . $index, $description );

    # Write the xdr:cNvPicPr element.
    $self->_write_c_nv_pic_pr();

    $self->{_writer}->endTag( 'xdr:nvPicPr' );
}


##############################################################################
#
# _write_c_nv_pic_pr()
#
# Write the <xdr:cNvPicPr> element.
#
sub _write_c_nv_pic_pr {

    my $self                 = shift;

    $self->{_writer}->startTag( 'xdr:cNvPicPr' );

    # Write the a:picLocks element.
    $self->_write_a_pic_locks();

    $self->{_writer}->endTag( 'xdr:cNvPicPr' );
}


##############################################################################
#
# _write_a_pic_locks()
#
# Write the <a:picLocks> element.
#
sub _write_a_pic_locks {

    my $self             = shift;
    my $no_change_aspect = 1;

    my @attributes = ( 'noChangeAspect' => $no_change_aspect );

    $self->{_writer}->emptyTag( 'a:picLocks', @attributes );
}


##############################################################################
#
# _write_blip_fill()
#
# Write the <xdr:blipFill> element.
#
sub _write_blip_fill {

    my $self  = shift;
    my $index = shift;

    $self->{_writer}->startTag( 'xdr:blipFill' );

    # Write the a:blip element.
    $self->_write_a_blip( $index );

    # Write the a:stretch element.
    $self->_write_a_stretch();

    $self->{_writer}->endTag( 'xdr:blipFill' );
}


##############################################################################
#
# _write_a_blip()
#
# Write the <a:blip> element.
#
sub _write_a_blip {

    my $self    = shift;
    my $index   = shift;
    my $schema  = 'http://schemas.openxmlformats.org/officeDocument/';
    my $xmlns_r = $schema . '2006/relationships';
    my $r_embed = 'rId' . $index;

    my @attributes = (
        'xmlns:r' => $xmlns_r,
        'r:embed' => $r_embed,
    );

    $self->{_writer}->emptyTag( 'a:blip', @attributes );
}


##############################################################################
#
# _write_a_stretch()
#
# Write the <a:stretch> element.
#
sub _write_a_stretch {

    my $self = shift;

    $self->{_writer}->startTag( 'a:stretch' );

    # Write the a:fillRect element.
    $self->_write_a_fill_rect();

    $self->{_writer}->endTag( 'a:stretch' );
}


##############################################################################
#
# _write_a_fill_rect()
#
# Write the <a:fillRect> element.
#
sub _write_a_fill_rect {

    my $self = shift;

    $self->{_writer}->emptyTag( 'a:fillRect' );
}


##############################################################################
#
# _write_sp_pr()
#
# Write the <xdr:spPr> element.
#
sub _write_sp_pr {

    my $self         = shift;
    my $col_absolute = shift;
    my $row_absolute = shift;
    my $width        = shift;
    my $height       = shift;

    $self->{_writer}->startTag( 'xdr:spPr' );

    # Write the a:xfrm element.
    $self->_write_a_xfrm( $col_absolute, $row_absolute, $width, $height );

    # Write the a:prstGeom element.
    $self->_write_a_prst_geom();

    $self->{_writer}->endTag( 'xdr:spPr' );
}


##############################################################################
#
# _write_a_xfrm()
#
# Write the <a:xfrm> element.
#
sub _write_a_xfrm {

    my $self         = shift;
    my $col_absolute = shift;
    my $row_absolute = shift;
    my $width        = shift;
    my $height       = shift;

    $self->{_writer}->startTag( 'a:xfrm' );

    # Write the a:off element.
    $self->_write_a_off( $col_absolute, $row_absolute );

    # Write the a:ext element.
    $self->_write_a_ext( $width, $height );

    $self->{_writer}->endTag( 'a:xfrm' );
}


##############################################################################
#
# _write_a_off()
#
# Write the <a:off> element.
#
sub _write_a_off {

    my $self = shift;
    my $x    = shift;
    my $y    = shift;

    my @attributes = (
        'x' => $x,
        'y' => $y,
    );

    $self->{_writer}->emptyTag( 'a:off', @attributes );
}


##############################################################################
#
# _write_a_ext()
#
# Write the <a:ext> element.
#
sub _write_a_ext {

    my $self = shift;
    my $cx   = shift;
    my $cy   = shift;

    my @attributes = (
        'cx' => $cx,
        'cy' => $cy,
    );

    $self->{_writer}->emptyTag( 'a:ext', @attributes );
}


##############################################################################
#
# _write_a_prst_geom()
#
# Write the <a:prstGeom> element.
#
sub _write_a_prst_geom {

    my $self = shift;
    my $prst = 'rect';

    my @attributes = ( 'prst' => $prst );

    $self->{_writer}->startTag( 'a:prstGeom', @attributes );

    # Write the a:avLst element.
    $self->_write_a_av_lst();

    $self->{_writer}->endTag( 'a:prstGeom' );
}


##############################################################################
#
# _write_a_av_lst()
#
# Write the <a:avLst> element.
#
sub _write_a_av_lst {

    my $self = shift;

    $self->{_writer}->emptyTag( 'a:avLst' );
}


1;


__END__

=pod

=head1 NAME

Drawing - A class for writing the Excel XLSX drawing.xml file.

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
