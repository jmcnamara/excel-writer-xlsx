package Excel::Writer::XLSX::Shape;

###############################################################################
#
# Shape - A class for writing Excel shapes.
#
# Used in conjunction with Excel::Writer::XLSX.
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
use Excel::Writer::XLSX::Worksheet;

our @ISA     = qw(Excel::Writer::XLSX::Worksheet);
our $VERSION = '0.46';

###############################################################################
#
# new()
#
sub new {

    my $class = shift;
    my $self  = Excel::Writer::XLSX::Package::XMLwriter->new();

    $self->{name} = undef;
    $self->{type} = 'rect';

    # isa Connector shape. 1/0 Value is a hash lookup from type
    $self->{connect} = 0;

    # isa Drawing, always 0, since a single shape never fills an entire sheet
    $self->{drawing} = 0;

    # OneCell or Absolute: options to move and/or size with cells
    $self->{editAs} = '';

    # Auto-incremented, unless supplied by user.
    $self->{id} = 0;

    $self->{text} = 0;

    # Index to _shapes array when inserted
    $self->{element} = -1;

    # Workbook format (for font, text decoration)
    $self->{format} = '';

    # Shape ID of starting connection, if any
    $self->{start} = undef;

    # Shape vertice, starts at 0, numbered clockwise from 12 oclock
    $self->{start_idx} = undef;

    $self->{end}     = undef;
    $self->{end_idx} = undef;

    # Number and size of adjustments for shapes (usually connectors)
    $self->{adjustments} = [];

    # t)op, b)ottom, l)eft, or r)ight
    $self->{start_side} = '';
    $self->{end_side}   = '';

    # Flip shape Horizontally. eg. arrow left to arrow right
    $self->{flipH} = 0;

    # Flip shape Vertically. eg. up arrow to down arrow
    $self->{flipV} = 0;

    # shape rotation (in degrees 0-360)
    $self->{rot} = 0;

    # Really just a rectangle with text
    $self->{txBox} = 0;

    # Shape outline color, or 0 for noFill (default black)
    $self->{line} = '000000';

    # dash, sysDot, dashDot, lgDash, lgDashDot, lgDashDotDot
    $self->{line_type} = '';

    # Line weight (integer)
    $self->{line_weight} = 1;

    # Shape fill color, or 0 for noFill (default noFill)
    $self->{fill} = 0;

    $self->{format}   = {};
    $self->{typeface} = 'Arial';

    # t, ctr, b
    $self->{valign} = 'ctr';

    # l, ctr, r, just
    $self->{align} = 'ctr';

    $self->{x_offset} = 0;
    $self->{y_offset} = 0;
    $self->{scale_x}  = 1;
    $self->{scale_y}  = 1;

    # Default size, which can be modified a/o scaled
    $self->{width}  = 50;
    $self->{height} = 50;

    # Initial assignment. May be modified when prepared
    $self->{column_start} = 0;
    $self->{row_start}    = 0;
    $self->{x1}           = 0;
    $self->{y1}           = 0;
    $self->{column_end}   = 0;
    $self->{row_end}      = 0;
    $self->{x2}           = 0;
    $self->{y2}           = 0;
    $self->{x_abs}        = 0;
    $self->{y_abs}        = 0;

    bless $self, $class;
    return $self;
}

1;

__END__

=head1 NAME

Shape - A class for creating Excel Drawing shapes

=head1 SYNOPSIS

To create a simple Excel file containing shapes using Excel::Writer::XLSX:

    #!/usr/bin/perl

    use strict;
    use warnings;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'shape.xlsx' );
    my $worksheet = $workbook->add_worksheet();

    # Add a default rectangle shape.
    my $rect = $workbook->add_shape();

    # Add an ellipse, with centered text.
    my $ellipse = $workbook->add_shape(
        type => 'ellipse',
        text => "Hello\nWorld"
    );

    # Add a cross, with a user-defined id.
    my $cross = $workbook->add_shape( type => 'cross', id => 33 );

    # Insert the shapes in the worksheet.
    $sheet->insert_shape( 'A1', $rect );
    $sheet->insert_shape( 'B2', $ellipse );
    $sheet->insert_shape( 'C3', $cross );

=head1 DESCRIPTION

The C<Excel::Writer::XLSX::Shape> module is used to create shape objects for L<Excel::Writer::XLSX>.

A shape object is created via the Workbook C<add_shape()> method:

    my $shape_rect = $workbook->add_shape( type => 'rect' );

Once the object is created it can be inserted into a worksheet using the C<insert_shape()> method. A shape can be inserted mulitple times if required.

    $sheet->insert_shape('A1', $shape_rect);
    $sheet->insert_shape('B2', $shape_rect, 20, 30);


=head1 METHODS

=head2 add_shape( %properties )

The C<add_shape()> Workboook method specifies the properties of the shape in hashC<( property => value )> format:

    my $shape = $workbook->add_shape( %proprties );

The available properties are shown below.


=head2 insert_shape( $row, $col, $shape, $x, $y, $scale_x, $scale_y )

The C<insert_shape()> Worksheet method sets the location and scale of the shape object within the worksheet.

    # Insert the shape into the a worksheet.
    $worksheet->insert_shape( 'E2', $shape );

Using the cell location and the C<$x> and C<$y> cell offsets it is possible to position a shape anywhere on the canvas of a worksheet.

A more detailed explanation of the C<insert_shape()> method is given in the main L<Excel::Writer::XLSX> documentation.


=head1 SHAPE PROPERTIES

The properties of a shape object that can be defined via C<add_shape()> are shown below.

=head2 name

Defines the name of the shape. This is optional.

=head2 type

Defines the type of the object such as C<rect>, C<ellipse> or C<triangle>:

    my $ellipse = $workbook->add_shape( type => 'ellipse' );

The full list of available shapes is shown below.

See also the C<all_shapes.pl> program in the C<examples> directory of the distro. It creates an example workbook with all the shapes, labelled with their shape names.

=over 4

=item * Basic Shapes

    rect parallelogram diamond roundRect octagon triangle rtTriangle ellipse
    hexagon plus pentagon can cube bevel foldedCorner smileyFace plaque
    leftBracket rightBracket leftBrace rightBrace

=item * Connectors

    line Bezier_curve path bezier_path straightConnector1 bentConnector3
    curvedConnector3

=item * Arrow Shapes

    rightArrow leftArrow upArrow downArrow leftRightArrow upDownArrow 4wayarrow
    3wayarrow curvedRightArrow curvedLeftArrow curvedUpArrow curvedDownArrow
    notchedRightArrow homePlate chevron rightArrowCallout leftArrowCallout
    upArrowCallout downArrowCallout leftRightArrowCallout upDownArrowCallout
    4wayarrowcallout

=item * Flow Chart Shapes

    flowChartProcess flowChartAlternateProcess flowChartDecision
    flowChartInputOutput flowChartPredefinedProcess flowChartInternalStorage
    flowChartDocument flowChartMultidocument flowChartTerminator
    flowChartPreparation flowChartManualInput flowChartManualOperation
    flowChartConnector flowChartOffpageConnector flowChartPunchedCard
    flowChartPunchedTape flowChartSummingJunction flowChartOr flowChartCollate
    flowChartSort flowChartExtract flowChartMerge flowChartOnlineStorage
    flowChartDelay flowChartMagneticTape flowChartMagneticDisk
    flowChartMagneticDrum flowChartDisplay

=item * Stars and Ribbons

    irregularSeal1 irregularSeal2 star4 star5 star8 star16 star24 star32 ribbon2
    ribbon ellipseRibbon2 ellipseRibbon verticalScroll horizontalScroll wave
    doubleWave

=item * Callout Shapes

    wedgeRectCallout wedgeRoundRectCallout wedgeEllipseCallout cloudCallout
    borderCallout1 borderCallout1 borderCallout2 borderCallout3 callout1
    accentCallout1 accentCallout2 accentCallout3 callout2 callout3 borderCallout1
    accentBorderCallout1 accentBorderCallout2 accentBorderCallout3

=back

=head2 text

This proprty is used to make the shape act like a text box.

    my $rect = $workbook->add_shape( type => 'rect', text => "Hello\nWorld" );

The text is super-imposed over the shape. The text can be wrapped using the newline character C<\n>.

=head2 id

Identification number for internal identification, or for identification in the resulting xml file. This number will be auto-assigned, if not assigned, or if it is a duplicate.

=head2 format

Workbook format for decorating shape text (font family, size, and decoration).

=head2 start, start_idx, end, end_idx

Shape ID of starting connection point for a connector, and index of connection. Index numbers are zero-based, and start from the top center, and count clockwise. Indices are are typically created for vertices and center points of shapes. They are the blue connection points that appear when connection shapes manually in Excel.

end and end_idx are for the connection end point, obviously.

=head2 start_side, end_side

This is either the letter C<b> or C<r> for the bottom or right side of the shape to be connected to and from.

If the start, start_idx, and start_side parameters are definded for a connection shape, the shape will be auto located and linked to the starting and ending shapes respectively. This can be very helpful for flow charts, organization charts, etc.

=head2 flipH, flipV

Set this value to 1, to flip the shape horizontally and/or vertically.

=head2 rot

Shape rotation, in degrees, from 0 to 360.

=head2 line, fill

Shape color for the outline and fill. Colors may be specified as a color index, or in rgb format, i.e. AA00FF.
see L<COLOURS IN EXCEL> for more information.

=head2 line_type

Line type for shape outline. The default is solid. The list of possible values is:

    dash, sysDot, dashDot, lgDash, lgDashDot, lgDashDotDot

=head2 valign, align

Text alignment within the shape. Vertical alignment may be t)top, ctr or b)ottom. Likewise,
horizontal alignment may be l, ctr, r, or just. The default is to center both horizontally and vertically.

=head2 scale_x, scale_y

Scale factor in x and y dimension, for scaling the shape width and height. Default value is 1.

=head2 adjustments

Adjustment of shape vertices. Most shapes do not use this. For some shapes, there is a single adjustment to modify the geometry. For instance, the plus shape has one adjustment to control the width of the spokes.

Connectors can have an odd number of adjustments to control the shape routing. Typically, a connector will have 3 or 5 handles for routing the shape. The adjustment is in percent of the distance from the starting shape to the ending shape, alternating between the x and y dimension. Adjustments may be negative, to route the shape away from the endpoint. The best way to learn about these is to play with them in Excel, and examine the xml that is produced.

The adjustment property must be supplied as an array reference [].


=head1 TIPS

Use C<< worksheet->hide_gridlines(2) >> to prepare a blank canvas without gridlines.

Shapes work in stencil mode. That is, once a shape is inserted, it is permanent, and can not be altered. The master shape may be modified after an instance is inserted, and only subsequent insertions will show the modifications. This is helpful for org charts, where a position shape may be created once, and then the name of the position modified for each position.

Shapes do not need to fit on one page. A large drawing may be created, and Excel will split the drawing into multiple pages. Use the page break preview to show page boundaries super imposed on the drawing.

Connected shapes will auto-locate in Excel, if you move either the starting shape or the ending shape separately. However, if you select both shapes (lasso or control-click), the connector will move with it, and the shape adjustments will not re-calculate.

=head1 EXAMPLE

A complete example is provided in the synopsis section. Also see shape1.pl, shape2.pl, ... and all_shapes.pl
in the examples folder of the distribution.

=head1 TO DO

=over 4

=item * Add shapes which have custom geometries

=item * Provide better integration of workbook formats for shapes

=item * Add validation of shape properties to prevent creation of workbooks that will not open.

=item * Auto connect shapes that are not anchored to cell A1

=item * Add automatic shape connection to shape vertices besides the object center.

=item * Improve automatic shape connection to shapes with concave sides (e.g. chevron).

=back

=head1 AUTHOR

Dave Clarke dclarke@cpan.org

=head1 COPYRIGHT

© MM-MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or
modified under the same terms as Perl itself.
