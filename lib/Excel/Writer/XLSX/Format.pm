package Excel::Writer::XLSX::Format;

###############################################################################
#
# Format - A class for defining Excel formatting.
#
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2010, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use 5.010000;
use Exporter;
use strict;
use warnings;
use Carp;


our @ISA     = qw(Exporter);
our $VERSION = '0.01';
our $AUTOLOAD;


###############################################################################
#
# new()
#
# Constructor
#
sub new {

    my $class = shift;

    my $self = {
        _xf_index => shift || 0,
        _palette => shift,

        _num_format     => 0,
        _font_index     => 0,
        _font           => 'Calibri',
        _size           => 11,
        _bold           => 0,
        _italic         => 0,
        _color          => 0x0,
        _underline      => 0,
        _font_strikeout => 0,
        _font_outline   => 0,
        _font_shadow    => 0,
        _font_script    => 0,
        _font_family    => 0,
        _font_charset   => 0,

        _hidden => 0,
        _locked => 1,

        _text_h_align  => 0,
        _text_wrap     => 0,
        _text_v_align  => -1,
        _text_justlast => 0,
        _rotation      => 0,
        _text_vertical => 0,

        _fg_color => 0x00,
        _bg_color => 0x00,

        _pattern => 0,

        _bottom => 0,
        _top    => 0,
        _left   => 0,
        _right  => 0,

        _bottom_color => 0x0,
        _top_color    => 0x0,
        _left_color   => 0x0,
        _right_color  => 0x0,

        _indent        => 0,
        _shrink        => 0,
        _merge_range   => 0,
        _reading_order => 0,

        _diag_type   => 0,
        _diag_color  => 0x0,
        _diag_border => 0,

        _just_distrib => 0,

    };

    bless $self, $class;

    # Set properties passed to Workbook::add_format()
    $self->set_format_properties(@_) if @_;

    return $self;
}


###############################################################################
#
# copy($format)
#
# Copy the attributes of another Excel::Writer::XLSX::Format object.
#
sub copy {
    my $self  = shift;
    my $other = $_[0];


    return unless defined $other;
    return unless ( ref( $self ) eq ref( $other ) );


    my $xf      = $self->{_xf_index};   # Store XF index assigned by Workbook.pm
    my $palette = $self->{_palette};    # Store palette assigned by Workbook.pm
    %$self             = %$other;       # Copy properties
    $self->{_xf_index} = $xf;           # Restore XF index
    $self->{_palette}  = $palette;      # Restore palette
}


###############################################################################
#
# convert_to_html_color()
#
# Convert from an Excel internal colour index to a Html style #RRGGBB index
# based on the default or user defined values in the Workbook palette.
#
sub convert_to_html_color {

    my $self  = shift;
    my $index = $_[0];

    return 0 unless $index;

    $index -= 8;    # Adjust colour index

    # _palette is a reference to the colour palette in the Workbook module
    my @rgb = @{ ${ $self->{_palette} }->[$index] }[ 0, 1, 2 ];

    return sprintf "#%02X%02X%02X", @rgb;
}


###############################################################################
#
# get_align_properties()
#
# Return properties for an Excel XML <Alignment> element.
#
# Excels handling of the vertical align "Bottom" property is different from
# other properties. It is on by default if any non-vertical property is set.
# Therefore we set the undefined _text_v_align value to -1 so that we can
# detect if it has been set by the user. If it hasn't been set then we supply
# the default "Bottom" value.
#
#
sub get_align_properties {

    my $self = shift;

    my @align;    # Attributes to return

    # Check if any alignment options in the format have been changed.
    my $changed =
      (      $self->{_text_h_align} != 0
          || $self->{_text_v_align} != -1
          || $self->{_indent} != 0
          || $self->{_rotation} != 0
          || $self->{_text_vertical} != 0
          || $self->{_text_wrap} != 0
          || $self->{_shrink} != 0
          || $self->{_reading_order} != 0 ) ? 1 : 0;


    return unless $changed;

    # Excel sets 'ss:Vertical="Bottom"' even when it is the default.
    $self->{_text_v_align} = 2 if $self->{_text_v_align} == -1;


    # Check for properties that are mutually exclusive.
    $self->{_rotation} = 0 if $self->{_text_vertical};
    $self->{_shrink}   = 0 if $self->{_text_wrap};
    $self->{_shrink}   = 0 if $self->{_text_h_align} == 4;    # Fill
    $self->{_shrink}   = 0 if $self->{_text_h_align} == 5;    # Justify
    $self->{_shrink}   = 0 if $self->{_text_h_align} == 7;    # Distributed
    $self->{_just_distrib} = 0
      if $self->{_text_h_align} != 7;                         # Distributed TODO


    push @align, 'ss:Horizontal', 'Left'    if $self->{_text_h_align} == 1;
    push @align, 'ss:Horizontal', 'Center'  if $self->{_text_h_align} == 2;
    push @align, 'ss:Horizontal', 'Right'   if $self->{_text_h_align} == 3;
    push @align, 'ss:Horizontal', 'Fill'    if $self->{_text_h_align} == 4;
    push @align, 'ss:Horizontal', 'Justify' if $self->{_text_h_align} == 5;
    push @align, 'ss:Horizontal', 'CenterAcrossSelection'
      if $self->{_text_h_align} == 6;
    push @align, 'ss:Horizontal', 'Distributed' if $self->{_text_h_align} == 7;

    push @align, 'ss:Vertical', 'Top'         if $self->{_text_v_align} == 0;
    push @align, 'ss:Vertical', 'Center'      if $self->{_text_v_align} == 1;
    push @align, 'ss:Vertical', 'Bottom'      if $self->{_text_v_align} == 2;
    push @align, 'ss:Vertical', 'Justify'     if $self->{_text_v_align} == 3;
    push @align, 'ss:Vertical', 'Distributed' if $self->{_text_v_align} == 4;

    push @align, 'ss:Indent', $self->{_indent}   if $self->{_indent};
    push @align, 'ss:Rotate', $self->{_rotation} if $self->{_rotation};

    push @align, 'ss:VerticalText', 1 if $self->{_text_vertical};
    push @align, 'ss:WrapText',     1 if $self->{_text_wrap};
    push @align, 'ss:ShrinkToFit',  1 if $self->{_shrink};

    # 'Context' is default property for ReadingOrder.
    push @align, 'ss:ReadingOrder', 'LeftToRight'
      if $self->{_reading_order} == 1;
    push @align, 'ss:ReadingOrder', 'RightToLeft'
      if $self->{_reading_order} == 2;


    # TODO
    #    ss:Horizontal="JustifyDistributed" ss:Vertical="Bottom"

    return @align;
}


###############################################################################
#
# get_border_properties()
#
# Return properties for an Excel XML <Border> element.
#
sub get_border_properties {

    my $self = shift;

    my @border;    # Attributes to return


    my %linetypes = (
        1  => [ 'ss:LineStyle' => 'Continuous',   'ss:Weight' => 1 ],
        2  => [ 'ss:LineStyle' => 'Continuous',   'ss:Weight' => 2 ],
        3  => [ 'ss:LineStyle' => 'Dash',         'ss:Weight' => 1 ],
        4  => [ 'ss:LineStyle' => 'Dot',          'ss:Weight' => 1 ],
        5  => [ 'ss:LineStyle' => 'Continuous',   'ss:Weight' => 3 ],
        6  => [ 'ss:LineStyle' => 'Double',       'ss:Weight' => 3 ],
        7  => [ 'ss:LineStyle' => 'Continuous' ],
        8  => [ 'ss:LineStyle' => 'Dash',         'ss:Weight' => 2 ],
        9  => [ 'ss:LineStyle' => 'DashDot',      'ss:Weight' => 1 ],
        10 => [ 'ss:LineStyle' => 'DashDot',      'ss:Weight' => 2 ],
        11 => [ 'ss:LineStyle' => 'DashDotDot',   'ss:Weight' => 1 ],
        12 => [ 'ss:LineStyle' => 'DashDotDot',   'ss:Weight' => 2 ],
        13 => [ 'ss:LineStyle' => 'SlantDashDot', 'ss:Weight' => 2 ],
    );


    for my $position ( '_bottom', '_left', '_right', '_top' ) {

        ( my $type = $position ) =~ s/^_//;
        my @attribs = ( 'ss:Position', ucfirst $type );
        my $position_color = $position . '_color';

        if ( exists $linetypes{ $self->{$position} } ) {

            push @attribs, @{ $linetypes{ $self->{$position} } };

            if ( my $color = $self->{$position_color} ) {
                $color = $self->convert_to_html_color( $color );
                push @attribs, 'ss:Color', $color;
            }

            push @border, [@attribs];
        }
    }


    # Handle diagonal borders. Note that in Excel it is only possible to have
    # one line type and one colour when both diagonals are in use.
    if ( my $diag_type = $self->{_diag_type} ) {

        # Set a default diagonal border style if none was specified.
        $self->{_diag_border} = 1 if not $self->{_diag_border};


        my @attribs = @{ $linetypes{ $self->{_diag_border} } };

        if ( my $color = $self->{_diag_color} ) {
            $color = $self->convert_to_html_color( $color );
            push @attribs, 'ss:Color', $color;
        }

        if ( $diag_type == 1 or $diag_type == 3 ) {
            push @border, [ "ss:Position", "DiagonalLeft", @attribs ];
        }

        if ( $diag_type == 2 or $diag_type == 3 ) {
            push @border, [ "ss:Position", "DiagonalRight", @attribs ];
        }
    }

    return @border;
}


###############################################################################
#
# get_font_properties()
#
# Return properties for an Excel XML <Font> element.
#
sub get_font_properties {

    my $self = shift;

    my @font;    # Attributes to return

    my $color = $self->convert_to_html_color( $self->{_color} );


    push @font, 'ss:FontName', $self->{_font} if $self->{_font} ne 'Arial';
    push @font, 'ss:Size',   $self->{_size} if $self->{_size} != 10;
    push @font, 'ss:Color',  $color         if $self->{_color};
    push @font, 'ss:Bold',   1              if $self->{_bold};
    push @font, 'ss:Italic', 1              if $self->{_italic};

    push @font, 'ss:StrikeThrough', 1 if $self->{_font_strikeout};
    push @font, 'ss:Outline',       1 if $self->{_font_outline};
    push @font, 'ss:Shadow',        1 if $self->{_font_shadow};

    push @font, 'ss:VerticalAlign', 'Superscript' if $self->{_font_script} == 1;
    push @font, 'ss:VerticalAlign', 'Subscript'   if $self->{_font_script} == 2;

    push @font, 'ss:Underline', 'Single'           if $self->{_underline} == 1;
    push @font, 'ss:Underline', 'Double'           if $self->{_underline} == 2;
    push @font, 'ss:Underline', 'SingleAccounting' if $self->{_underline} == 33;
    push @font, 'ss:Underline', 'DoubleAccounting' if $self->{_underline} == 34;

    push @font, 'x:Family',  $self->{_font_family}  if $self->{_font_family};
    push @font, 'x:CharSet', $self->{_font_charset} if $self->{_font_charset};

    return @font;
}


###############################################################################
#
# get_interior_properties()
#
# Return properties for an Excel XML <Interior> element.
#
sub get_interior_properties {

    my $self = shift;

    # Return undef if the background and foreground colours haven't been set
    # and the pattern hasn't been set or if it has only been set to solid.
    # Other patterns will be handled with the default colours.
    #
    return
      if $self->{_fg_color} == 0x00
          and $self->{_bg_color} == 0x00
          and $self->{_pattern} <= 0x01;


    # Note for XML:
    #               ss:Color        = _bg_color
    #               ss:PatternColor = _fg_color


    # The following logical statements take care of special cases in relation
    # to cell colours and patterns:
    # 1. For a solid fill (_pattern == 1) Excel reverses the role of foreground
    #    and background colours.
    # 2. If the user specifies a foreground or background colour without a
    #    pattern they probably wanted a solid fill, so we fill in the defaults.
    #
    if ( $self->{_pattern} <= 0x01 ) {
        if ( $self->{_bg_color} ) {
            return 'ss:Color',
              $self->convert_to_html_color( $self->{_bg_color} ),
              'ss:Pattern',
              'Solid';
        }
        else {
            return 'ss:Color',
              $self->convert_to_html_color( $self->{_fg_color} ),
              'ss:Pattern',
              'Solid';
        }
    }


    # Set default colours if they haven't been set.
    $self->{_bg_color} = 0x09 if $self->{_bg_color} == 0x00;    # 0x09 = white
    $self->{_fg_color} = 0x08 if $self->{_fg_color} == 0x00;    # 0x08 = black

    my %patterns = (
        1  => 'Solid',
        2  => 'Gray50',
        3  => 'Gray75',
        4  => 'Gray25',
        5  => 'HorzStripe',
        6  => 'VertStripe',
        7  => 'ReverseDiagStripe',
        8  => 'DiagStripe',
        9  => 'DiagCross',
        10 => 'ThickDiagCross',
        11 => 'ThinHorzStripe',
        12 => 'ThinVertStripe',
        13 => 'ThinReverseDiagStripe',
        14 => 'ThinDiagStripe',
        15 => 'ThinHorzCross',
        16 => 'ThinDiagCross',
        17 => 'Gray125',
        18 => 'Gray0625',
    );

    return unless exists $patterns{ $self->{_pattern} };

    return 'ss:Color',
      $self->convert_to_html_color( $self->{_bg_color} ),
      'ss:Pattern',
      $patterns{ $self->{_pattern} },
      'ss:PatternColor',
      $self->convert_to_html_color( $self->{_fg_color} );
}


###############################################################################
#
# get_num_format_properties()
#
# Return properties for an Excel XML <NumberFormat> element.
#
sub get_num_format_properties {

    my $self = shift;

    return unless defined $self->{_num_format};


    # This hash is here mainly to cater for Spreadsheet::WriteExcel programs
    # and Excel files that use the in-built format codes. ExcelXML users
    # should specify the format explicitly.
    #
    my %num_format = (
        1  => '0',
        2  => 'Fixed',
        3  => '#,##0',
        4  => 'Standard',
        5  => '$#,##0;\-$#,##0',
        6  => '$#,##0;[Red]\-$#,##0',
        7  => '$#,##0.00;\-$#,##0.00',
        8  => 'Currency',
        9  => '0%',
        10 => 'Percent',
        11 => 'Scientific',
        12 => '#\ ?/?',
        13 => '#\ ??/??',
        14 => 'Short Date',
        15 => 'Medium Date',
        16 => 'dd\-mmm',
        17 => 'mmm\-yy',
        18 => 'Medium Time',
        19 => 'Long Time',
        20 => 'Short Time',
        21 => 'hh:mm:ss',
        22 => 'General Date',
        37 => '#,##0;\-#,##0',
        38 => '#,##0;[Red]\-#,##0',
        39 => '#,##0.00;\-#,##0.00',
        40 => '#,##0.00;[Red]\-#,##0.00',
        41 => '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-',
        42 => '_-$* #,##0_-;\-$* #,##0_-;_-$* "-"_-;_-@_-',
        43 => '_-* #,##0.00_-;\-* #,##0.00_-;_-* "-"??_-;_-@_-',
        44 => '_-$* #,##0.00_-;\-$* #,##0.00_-;_-$* "-"??_-;_-@_-',
        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mm:ss.0',
        48 => '##0.0E+0',
        49 => '@',
    );

    my $num_format;

    # Num_format is either a built-in code or a user specified string.
    if ( exists $num_format{ $self->{_num_format} } ) {
        $num_format = $num_format{ $self->{_num_format} };
    }
    else {
        $num_format = $self->{_num_format};
    }

    return 'ss:Format', $num_format;
}


###############################################################################
#
# get_protection_properties()
#
# Return properties for an Excel XML <Protection> element.
#
sub get_protection_properties {

    my $self = shift;

    my @attribs;    # Attributes to return

    push @attribs, 'x:HideFormula', 1 if $self->{_hidden};
    push @attribs, 'ss:Protected',  0 if not $self->{_locked};

    return @attribs;
}


###############################################################################
#
# get_font_key()
#
# Returns a unique hash key for a font. Used by Workbook->_store_all_fonts()
#
sub get_font_key {

    my $self    = shift;

    # The following elements are arranged to increase the probability of
    # generating a unique key. Elements that hold a large range of numbers
    # e.g. _color are placed between two binary elements such as _italic
    #
    my $key = "$self->{_font}$self->{_size}";
    $key   .= "$self->{_font_script}$self->{_underline}";
    $key   .= "$self->{_font_strikeout}$self->{_bold}$self->{_font_outline}";
    $key   .= "$self->{_font_family}$self->{_font_charset}";
    $key   .= "$self->{_font_shadow}$self->{_color}$self->{_italic}";
    $key    =~ s/ /_/g; # Convert the key to a single word

    return $key;
}


###############################################################################
#
# get_xf_index()
#
# Returns the index used by Worksheet->_XF()
#
sub get_xf_index {
    my $self = shift;

    return $self->{_xf_index};
}


###############################################################################
#
# _get_color()
#
# Used in conjunction with the set_xxx_color methods to convert a color
# string into a number. Color range is 0..63 but we will restrict it
# to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
#
sub _get_color {

    my %colors = (
        aqua    => 0x0F,
        cyan    => 0x0F,
        black   => 0x08,
        blue    => 0x0C,
        brown   => 0x10,
        magenta => 0x0E,
        fuchsia => 0x0E,
        gray    => 0x17,
        grey    => 0x17,
        green   => 0x11,
        lime    => 0x0B,
        navy    => 0x12,
        orange  => 0x35,
                    pink    => 0x21,
        purple  => 0x14,
        red     => 0x0A,
        silver  => 0x16,
        white   => 0x09,
        yellow  => 0x0D,
    );

    # Return the default color if undef,
    return 0x00 unless defined $_[0];

    # or the color string converted to an integer,
    return $colors{ lc( $_[0] ) } if exists $colors{ lc( $_[0] ) };

    # or the default color if string is unrecognised,
    return 0x00 if ( $_[0] =~ m/\D/ );

    # or an index < 8 mapped into the correct range,
    return $_[0] + 8 if $_[0] < 8;

    # or the default color if arg is outside range,
    return 0x00 if $_[0] > 63;

    # or an integer in the valid range
    return $_[0];
}


###############################################################################
#
# set_type()
#
# Set the XF object type as 0 = cell XF or 0xFFF5 = style XF.
#
sub set_type {

    my $self = shift;
    my $type = $_[0];

    if (defined $_[0] and $_[0] eq 0) {
        $self->{_type} = 0x0000;
    }
    else {
        $self->{_type} = 0xFFF5;
    }
}


###############################################################################
#
# set_align()
#
# Set cell alignment.
#
sub set_align {

    my $self     = shift;
    my $location = $_[0];

    return if not defined $location;    # No default
    return if $location =~ m/\d/;       # Ignore numbers

    $location = lc( $location );

    $self->set_text_h_align( 1 ) if ( $location eq 'left' );
    $self->set_text_h_align( 2 ) if ( $location eq 'centre' );
    $self->set_text_h_align( 2 ) if ( $location eq 'center' );
    $self->set_text_h_align( 3 ) if ( $location eq 'right' );
    $self->set_text_h_align( 4 ) if ( $location eq 'fill' );
    $self->set_text_h_align( 5 ) if ( $location eq 'justify' );
    $self->set_text_h_align( 6 ) if ( $location eq 'center_across' );
    $self->set_text_h_align( 6 ) if ( $location eq 'centre_across' );
    $self->set_text_h_align( 6 ) if ( $location eq 'merge' );        # S:WE name
    $self->set_text_h_align( 7 ) if ( $location eq 'distributed' );
    $self->set_text_h_align( 7 ) if ( $location eq 'equal_space' ); # ParseExcel


    $self->set_text_v_align( 0 ) if ( $location eq 'top' );
    $self->set_text_v_align( 1 ) if ( $location eq 'vcentre' );
    $self->set_text_v_align( 1 ) if ( $location eq 'vcenter' );
    $self->set_text_v_align( 2 ) if ( $location eq 'bottom' );
    $self->set_text_v_align( 3 ) if ( $location eq 'vjustify' );
    $self->set_text_v_align( 4 ) if ( $location eq 'vdistributed' );
    $self->set_text_v_align( 4 )
      if ( $location eq 'vequal_space' );                           # ParseExcel
}


###############################################################################
#
# set_valign()
#
# Set vertical cell alignment. This is required by the set_properties() method
# to differentiate between the vertical and horizontal properties.
#
sub set_valign {

    my $self = shift;
    $self->set_align( @_ );
}


###############################################################################
#
# set_center_across()
#
# Implements the Excel5 style "merge".
#
sub set_center_across {

    my $self = shift;

    $self->set_text_h_align( 6 );
}


###############################################################################
#
# set_merge()
#
# This was the way to implement a merge in Excel5. However it should have been
# called "center_across" and not "merge".
# This is now deprecated. Use set_center_across() or better merge_range().
#
#
sub set_merge {

    my $self = shift;

    $self->set_text_h_align( 6 );
}


###############################################################################
#
# set_bold()
#
# Unlike the binary format in Spreadsheet::WriteExcel bold cannot have a
# "weight". In the XML format it is either on or off.
#
sub set_bold {

    my $self = shift;
    $self->{_bold} = $_[0] ? 1 : 0;
}


###############################################################################
#
# set_border($style)
#
# Set cells borders to the same style
#
sub set_border {

    my $self  = shift;
    my $style = $_[0];

    $self->set_bottom( $style );
    $self->set_top( $style );
    $self->set_left( $style );
    $self->set_right( $style );
}


###############################################################################
#
# set_border_color($color)
#
# Set cells border to the same color
#
sub set_border_color {

    my $self  = shift;
    my $color = $_[0];

    $self->set_bottom_color( $color );
    $self->set_top_color( $color );
    $self->set_left_color( $color );
    $self->set_right_color( $color );
}


###############################################################################
#
# set_rotation($angle)
#
# Set the rotation angle of the text. An alignment property.
#
sub set_rotation {

    my $self     = shift;
    my $rotation = $_[0];

    # Argument should be a number
    return if $rotation !~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/;

    # The arg type can be a double but the Excel dialog only allows integers.
    $rotation = int $rotation;

    if ( $rotation == 270 ) {

        # Special case inherited from the S::WE interface.
        $self->{_text_vertical} = 1;
        $self->{_rotation}      = 0;
        return;
    }
    elsif ( $rotation < -90 or $rotation > 90 ) {
        carp "Rotation $rotation outside range: -90 <= angle <= 90";
        $self->{_rotation} = 0;
        return;
    }

    # Rotation and vertical text are mutually exclusive
    $self->{_text_vertical} = 0;
    $self->{_rotation}      = $rotation;
}


###############################################################################
#
# set_format_properties()
#
# Convert hashes of properties to method calls.
#
sub set_format_properties {

    my $self = shift;

    my %properties = @_;    # Merge multiple hashes into one

    while ( my ( $key, $value ) = each( %properties ) ) {

        # Strip leading "-" from Tk style properties e.g. -color => 'red'.
        $key =~ s/^-//;

        # Create a sub to set the property.
        my $sub = \&{"set_$key"};
        $sub->($self, $value);
        }
        }

# Renamed rarely used set_properties() to set_format_properties() to avoid
# confusion with Workbook method of the same name. The following acts as an
# alias for any code that uses the old name.
*set_properties = *set_format_properties;


###############################################################################
#
# AUTOLOAD. Deus ex machina.
#
# Dynamically create set methods that aren't already defined.
#
sub AUTOLOAD {

    my $self = shift;

    # Ignore calls to DESTROY
    return if $AUTOLOAD =~ /::DESTROY$/;

    # Check for a valid method names, i.e. "set_xxx_yyy".
    $AUTOLOAD =~ /.*::set(\w+)/ or die "Unknown method: $AUTOLOAD\n";

    # Match the attribute, i.e. "_xxx_yyy".
    my $attribute = $1;

    # Check that the attribute exists
    exists $self->{$attribute} or die "Unknown method: $AUTOLOAD\n";

    # The attribute value
    my $value;


    # There are two types of set methods: set_property() and
    # set_property_color(). When a method is AUTOLOADED we store a new anonymous
    # sub in the appropriate slot in the symbol table. The speeds up subsequent
    # calls to the same method.
    #
    no strict 'refs';    # To allow symbol table hackery

    if ( $AUTOLOAD =~ /.*::set\w+color$/ ) {

        # For "set_property_color" methods
        $value = _get_color( $_[0] );

        *{$AUTOLOAD} = sub {
            my $self = shift;

            $self->{$attribute} = _get_color( $_[0] );
        };
    }
    else {

        $value = $_[0];
        $value = 1 if not defined $value;    # The default value is always 1

        *{$AUTOLOAD} = sub {
            my $self  = shift;
            my $value = shift;

            $value = 1 if not defined $value;
            $self->{$attribute} = $value;
        };
    }


    $self->{$attribute} = $value;
}


1;


__END__


=head1 NAME

Format - A class for defining Excel formatting.

=head1 SYNOPSIS

See the documentation for Excel::Writer::XLSX

=head1 DESCRIPTION

This module is used in conjunction with Excel::Writer::XLSX.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

� MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
