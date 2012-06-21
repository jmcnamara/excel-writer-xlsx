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
our $VERSION = '0.47';
our $AUTOLOAD;

###############################################################################
#
# new()
#
sub new {

    my $class = shift;
    my %properties = @_;
    my $self  = Excel::Writer::XLSX::Package::XMLwriter->new();

    $self->{_name} = undef;
    $self->{_type} = 'rect';

    # isa Connector shape. 1/0 Value is a hash lookup from type
    $self->{_connect} = 0;

    # isa Drawing, always 0, since a single shape never fills an entire sheet
    $self->{_drawing} = 0;

    # OneCell or Absolute: options to move and/or size with cells
    $self->{_editAs} = '';

    # Auto-incremented, unless supplied by user.
    $self->{_id} = 0;

    # Shape text (usually centered on shape geometry)
    $self->{_text} = 0;

    # Shape stencil mode.  A copy (child) is created when inserted.  Link to parent is broken.
    $self->{_stencil} = 1;

    # Index to _shapes array when inserted
    $self->{_element} = -1;

    # Shape ID of starting connection, if any
    $self->{_start} = undef;

    # Shape vertice, starts at 0, numbered clockwise from 12 oclock
    $self->{_start_index} = undef;

    $self->{_end}     = undef;
    $self->{_end_index} = undef;

    # Number and size of adjustments for shapes (usually connectors)
    $self->{_adjustments} = [];

    # t)op, b)ottom, l)eft, or r)ight
    $self->{_start_side} = '';
    $self->{_end_side}   = '';

    # Flip shape Horizontally. eg. arrow left to arrow right
    $self->{_flip_h} = 0;

    # Flip shape Vertically. eg. up arrow to down arrow
    $self->{_flip_v} = 0;

    # shape rotation (in degrees 0-360)
    $self->{_rotation} = 0;

    # An alternate way to create a text box, because Excel allows it.  It is just a rectangle with text
    $self->{_txBox} = 0;

    # Shape outline color, or 0 for noFill (default black)
    $self->{_line} = '000000';

    # dash, sysDot, dashDot, lgDash, lgDashDot, lgDashDotDot
    $self->{_line_type} = '';

    # Line weight (integer)
    $self->{_line_weight} = 1;

    # Shape fill color, or 0 for noFill (default noFill)
    $self->{_fill} = 0;

    # Formatting for shape text, if any
    $self->{_format}   = {};

    # copy of color palette table from Workbook.pm
    $self->{_palette}   = [];

    # t, ctr, b
    $self->{_valign} = 'ctr';

    # l, ctr, r, just
    $self->{_align} = 'ctr';

    $self->{_x_offset} = 0;
    $self->{_y_offset} = 0;

    # Scale factors, which also may be set when the shape is inserted.
    $self->{_scale_x}  = 1;
    $self->{_scale_y}  = 1;

    # Default size, which can be modified a/o scaled
    $self->{_width}  = 50;
    $self->{_height} = 50;

    # Initial assignment. May be modified when prepared
    $self->{_column_start} = 0;
    $self->{_row_start}    = 0;
    $self->{_x1}           = 0;
    $self->{_y1}           = 0;
    $self->{_column_end}   = 0;
    $self->{_row_end}      = 0;
    $self->{_x2}           = 0;
    $self->{_y2}           = 0;
    $self->{_x_abs}        = 0;
    $self->{_y_abs}        = 0;

    # Override default properties with passed arguments
    while ( my ( $key, $value ) = each( %properties ) ) {

        # Strip leading "-" from Tk style properties e.g. -color => 'red'.
        $key =~ s/^-//;

        # Add leading underscore "_" to internal hash keys, if not supplied.
        $key = "_" . $key unless $key =~ m/^_/;

        $self->{$key} = $value;
    }

    bless $self, $class;
    return $self;
}
###############################################################################
#
# set_properties ( name => 'Shape 1', type => 'rect' )
#
# Set shape properties 
#
sub set_properties {
    my $self = shift;
    my %properties = @_;
    # Update properties with passed arguments
    while ( my ( $key, $value ) = each( %properties ) ) {

        # Strip leading "-" from Tk style properties e.g. -color => 'red'.
        $key =~ s/^-//;

        # Add leading underscore "_" to internal hash keys, if not supplied.
        $key = "_" . $key unless $key =~ m/^_/;

        exists $self->{$key} or do {warn "Unknown shape property: $key.  Property not set\n"; next};
        $self->{$key} = $value;
    }
}

###############################################################################
#
# set_adjustment ( adj1, adj2, adj3, ... )
#
# Set the shape adjustments array (as a reference)
#
sub set_adjustments {
    my $self = shift;
    $self->{_adjustments} = \@_;
}

###############################################################################
#
# AUTOLOAD. Deus ex machina.
#
# Dynamically create set/get methods that aren't already defined.
#
sub AUTOLOAD {

    my $self = shift;

    # Ignore calls to DESTROY
    return if $AUTOLOAD =~ /::DESTROY$/;

    # Check for a valid method names, i.e. "set_xxx_yyy".
    $AUTOLOAD =~ /.*::(get|set)(\w+)/ or die "Unknown method: $AUTOLOAD\n";

    # Match the function (get or set) and attribute, i.e. "_xxx_yyy".
    my $gs = $1;
    my $attribute = $2;

    # Check that the attribute exists
    exists $self->{$attribute} or die "Unknown method: $AUTOLOAD\n";

    # The attribute value
    my $value;

    # set_property() pattern
    # When a method is AUTOLOADED we store a new anonymous
    # sub in the appropriate slot in the symbol table. The speeds up subsequent
    # calls to the same method.
    #
    no strict 'refs';    # To allow symbol table hackery

    $value = $_[0];
    $value = 1 if not defined $value;    # The default value is always 1

    if ($gs eq 'set') {
        *{$AUTOLOAD} = sub {
            my $self  = shift;
            my $value = shift;
    
            $value = 1 if not defined $value;
            $self->{$attribute} = $value;
        };

        $self->{$attribute} = $value;
    } else {
        *{$AUTOLOAD} = sub {
            my $self  = shift;
            return $self->{$attribute};
        };

        # Let AUTOLOAD return the attribute for the first invocation
        return $self->{$attribute};
    }
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

Any shape property can be queried/modifed by the corresponding get/set method:
    
    my $ellipse = $workbook->add_shape( %proprties );
    $ellipse->set_type('cross');            # No longer an ellipse!
    my $type = $elipse->get_type();         # Find out what it really is

Multiple shape properties may also be modifed, by using the C<set_shape_properties> method.

        $shape->set_properties( type => 'ellipse', text => 'A Circle' );

The properties of a shape object that can be defined via C<add_shape()> are shown below.

=head2 name

Defines the name of the shape. This is optional.

=head2 type

Defines the type of the object such as C<rect>, C<ellipse> or C<triangle>:

    my $ellipse = $workbook->add_shape( type => 'ellipse' );
    

The full list of available shapes is shown below.

See also the C<all_shapes.pl> program in the C<examples> directory of the distro.
It creates an example workbook with all the shapes, labelled with their shape names.

The list in the consists of all the shape types
defined under xsd:simpleType name="ST_ShapeType" in ECMA-376 Office Open XML File Formats Part 4
the grouping by tab name is not part of the standard.

=over 4

=item * Action Shapes

actionButtonBackPrevious actionButtonBeginning actionButtonBlank 
actionButtonDocument actionButtonEnd actionButtonForwardNext 
actionButtonHelp actionButtonHome actionButtonInformation actionButtonMovie 
actionButtonReturn actionButtonSound

=item * Arrow Shapes

bentArrow bentUpArrow circularArrow curvedDownArrow curvedLeftArrow 
curvedRightArrow curvedUpArrow downArrow leftArrow leftCircularArrow 
leftRightArrow leftRightCircularArrow leftRightUpArrow leftUpArrow 
notchedRightArrow quadArrow rightArrow stripedRightArrow swooshArrow upArrow 
upDownArrow uturnArrow

=item * Basic Shapes

blockArc can chevron cube decagon diamond dodecagon donut ellipse funnel 
gear6 gear9 heart heptagon hexagon homePlate lightningBolt line lineInv moon 
nonIsoscelesTrapezoid noSmoking octagon parallelogram pentagon pie pieWedge 
plaque rect round1Rect round2DiagRect round2SameRect roundRect rtTriangle 
smileyFace snip1Rect snip2DiagRect snip2SameRect snipRoundRect star10 star12 
star16 star24 star32 star4 star5 star6 star7 star8 sun teardrop trapezoid 
triangle

=item * Callout Shapes

accentBorderCallout1 accentBorderCallout2 accentBorderCallout3 
accentCallout1 accentCallout2 accentCallout3 borderCallout1 borderCallout2 
borderCallout3 callout1 callout2 callout3 cloudCallout downArrowCallout 
leftArrowCallout leftRightArrowCallout quadArrowCallout rightArrowCallout 
upArrowCallout upDownArrowCallout wedgeEllipseCallout wedgeRectCallout 
wedgeRoundRectCallout

=item * Chart Shapes 

Not to be confused with Excel Charts.  There is no relationship.

chartPlus chartStar chartX

=item * Connector Shapes

bentConnector2 bentConnector3 bentConnector4 bentConnector5 curvedConnector2 
curvedConnector3 curvedConnector4 curvedConnector5 straightConnector1

=item * Arrow Shapes

rightArrow leftArrow upArrow downArrow leftRightArrow upDownArrow 4wayarrow
3wayarrow curvedRightArrow curvedLeftArrow curvedUpArrow curvedDownArrow
notchedRightArrow homePlate chevron rightArrowCallout leftArrowCallout
upArrowCallout downArrowCallout leftRightArrowCallout upDownArrowCallout
4wayarrowcallout

=item * Flow Chart Shapes

flowChartAlternateProcess   flowChartCollate            flowChartConnector 
flowChartDecision           flowChartDelay              flowChartDisplay 
flowChartDocument           flowChartExtract            flowChartInputOutput 
flowChartInternalStorage    flowChartMagneticDisk       flowChartMagneticDrum 
flowChartMagneticTape       flowChartManualInput        flowChartManualOperation 
flowChartMerge              flowChartMultidocument      flowChartOfflineStorage 
flowChartOffpageConnector   flowChartOnlineStorage      flowChartOr 
flowChartPredefinedProcess  flowChartPreparation        flowChartProcess 
flowChartPunchedCard        flowChartPunchedTape        flowChartSort 
flowChartSummingJunction    flowChartTerminator         

=item * Math Shapes

mathDivide   mathEqual    mathMinus    mathMultiply mathNotEqual mathPlus     

=item * Stars and Banners

arc             bevel           bracePair       bracketPair     chord 
cloud           corner          diagStripe      doubleWave      ellipseRibbon 
ellipseRibbon2  foldedCorner    frame           halfFrame       horizontalScroll 
irregularSeal1  irregularSeal2  leftBrace       leftBracket     leftRightRibbon 
plus            ribbon          ribbon2         rightBrace      rightBracket 
verticalScroll  wave            

=item * Tab Shapes

cornerTabs plaqueTabs squareTabs

=back

=head2 text

This property is used to make the shape act like a text box.

    my $rect = $workbook->add_shape( type => 'rect', text => "Hello\nWorld" );

The text is super-imposed over the shape. The text can be wrapped using the newline character C<\n>.

=head2 id

Identification number for internal identification, or for identification in the resulting xml file. This number will be auto-assigned, if not assigned, or if it is a duplicate.

=head2 format

Workbook format for decorating shape text (font family, size, and decoration).

=head2 start, start_index, end, end_index

Shape ID of starting connection point for a connector, and index of connection. Index numbers are zero-based, and start from the top center, and count clockwise. Indices are are typically created for vertices and center points of shapes. They are the blue connection points that appear when connection shapes manually in Excel.

end and end_index are for the connection end point, obviously.

=head2 start_side, end_side

This is either the letter C<b> or C<r> for the bottom or right side of the shape to be connected to and from.

If the start, start_index, and start_side parameters are defined for a connection shape, the shape will be auto located and linked to the starting and ending shapes respectively. This can be very helpful for flow charts, organization charts, etc.

=head2 flip_h, flip_v

Set this value to 1, to flip the shape horizontally and/or vertically.

=head2 rotation

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
Scaling may be set on the shape object, or adjusted via C<< insert_shape >>.

=head2 adjustments

Adjustment of shape vertices. Most shapes do not use this. For some shapes, there is a single adjustment to modify the geometry. For instance, the plus shape has one adjustment to control the width of the spokes.

Connectors can have a number of adjustments to control the shape routing. Typically, a connector will have 3 to 5 handles for routing the shape. The adjustment is in percent of the distance from the starting shape to the ending shape, alternating between the x and y dimension. Adjustments may be negative, to route the shape away from the endpoint. The best way to learn about these is to play with them in Excel, and examine the xml that is produced.

=head2 stencil

Shapes work in stencil mode by default. That is, once a shape is inserted, its connection
is separated from its master. The master shape may be modified after an 
instance is inserted, and only subsequent insertions will show the 
modifications. This is helpful for org charts, where an employee shape may be 
created once, and then the text of the shape is modified for each employee.
the C<< insert_shape >> method returns a reference to the inserted shape (the child).

Stencil mode can be turned off, allowing for shape(s) to be modified after insertion.
the C<< insert_shape >> method returns a reference to the inserted shape (the master).
This is not very useful for inserting multiple shapes, since the x/y coordinates also 
get modified.

=head1 TIPS

Use C<< worksheet->hide_gridlines(2) >> to prepare a blank canvas without gridlines.

Shapes do not need to fit on one page. A large drawing may be created, and Excel 
will split the drawing into multiple pages. Use the page break preview to show 
page boundaries super imposed on the drawing.

Connected shapes will auto-locate in Excel, if you move either the starting 
shape or the ending shape separately. However, if you select both shapes (lasso 
or control-click), the connector will move with it, and the shape adjustments 
will not re-calculate.

=head1 EXAMPLE

A complete example is provided in the synopsis section. Also see shape1.pl, shape2.pl, ... and shape_all.pl
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
