package Excel::Writer::XLSX::Chart;

###############################################################################
#
# Chart - A class for writing Excel Charts.
#
#
# Used in conjunction with Excel::Writer::XLSX.
#
# Copyright 2000-2011, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.010000;
use strict;
use warnings;
use Carp;
use Excel::Writer::XLSX::Format;
use Excel::Writer::XLSX::Package::XMLwriter;
use Excel::Writer::XLSX::Utility qw(xl_cell_to_rowcol
  xl_rowcol_to_cell
  xl_col_to_name xl_range
  xl_range_formula );

our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
our $VERSION = '0.18';


###############################################################################
#
# factory()
#
# Factory method for returning chart objects based on their class type.
#
sub factory {

    my $current_class  = shift;
    my $chart_subclass = shift;

    $chart_subclass = ucfirst lc $chart_subclass;

    my $module = "Excel::Writer::XLSX::Chart::" . $chart_subclass;

    eval "require $module";

    # TODO. Need to re-raise this error from Workbook::add_chart().
    die "Chart type '$chart_subclass' not supported in add_chart()\n" if $@;

    return $module->new( @_ );
}


###############################################################################
#
# new()
#
# Default constructor for sub-classes.
#
sub new {

    my $class = shift;
    my $self  = Excel::Writer::XLSX::Package::XMLwriter->new();

    $self->{_subtype}           = shift;
    $self->{_sheet_type}        = 0x0200;
    $self->{_orientation}       = 0x0;
    $self->{_series}            = [];
    $self->{_embedded}          = 0;
    $self->{_id}                = '';
    $self->{_style_id}          = 2;
    $self->{_axis_ids}          = [];
    $self->{_has_category}      = 0;
    $self->{_requires_category} = 0;
    $self->{_legend_position}   = 'right';
    $self->{_cat_axis_position} = 'b';
    $self->{_val_axis_position} = 'l';
    $self->{_formula_ids}       = {};
    $self->{_formula_data}      = [];
    $self->{_horiz_cat_axis}    = 0;
    $self->{_horiz_val_axis}    = 1;
    $self->{_protection}        = 0;

    bless $self, $class;
    $self->_set_default_properties();
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

    $self->_write_xml_declaration();


    # Write the c:chartSpace element.
    $self->_write_chart_space();

    # Write the c:lang element.
    $self->_write_lang();

    # Write the c:style element.
    $self->_write_style();

    # Write the c:protection element.
    $self->_write_protection();

    # Write the c:chart element.
    $self->_write_chart();

    # Write the c:printSettings element.
    $self->_write_print_settings() if $self->{_embedded};;

    # Close the worksheet tag.
    $self->{_writer}->endTag( 'c:chartSpace' );

    # Close the XML writer object and filehandle.
    $self->{_writer}->end();
    $self->{_writer}->getOutput()->close();
}


###############################################################################
#
# Public methods.
#
###############################################################################


###############################################################################
#
# add_series()
#
# Add a series and it's properties to a chart.
#
sub add_series {

    my $self = shift;
    my %arg  = @_;

    # Check that the required input has been specified.
    if ( !exists $arg{values} ) {
        croak "Must specify 'values' in add_series()";
    }

    if ( $self->{_requires_category} && !exists $arg{categories} ) {
        croak "Must specify 'categories' in add_series() for this chart type";
    }


    # Convert aref params into a formula string.
    my $values       = $self->_aref_to_formula( $arg{values} );
    my $categories   = $self->_aref_to_formula( $arg{categories} );


    # Switch name and name_formula parameters if required.
    my ( $name, $name_formula ) =
      $self->_process_names( $arg{name}, $arg{name_formula} );

    # Get an id for the data equivalent to the range formula.
    my $cat_id  = $self->_get_data_id( $categories,   $arg{categories_data} );
    my $val_id  = $self->_get_data_id( $values,       $arg{values_data} );
    my $name_id = $self->_get_data_id( $name_formula, $arg{name_data} );

    # Add the parsed data to the user supplied data. TODO. Refactor.
    %arg = (
        _values       => $values,
        _categories   => $categories,
        _name         => $name,
        _name_formula => $name_formula,
        _name_id      => $name_id,
        _val_data_id  => $val_id,
        _cat_data_id  => $cat_id,
    );

    push @{ $self->{_series} }, \%arg;
}


###############################################################################
#
# set_x_axis()
#
# Set the properties of the X-axis.
#
sub set_x_axis {

    my $self = shift;
    my %arg  = @_;

    my ( $name, $name_formula ) =
      $self->_process_names( $arg{name}, $arg{name_formula} );

    my $data_id = $self->_get_data_id( $name_formula, $arg{data} );

    $self->{_x_axis_name}    = $name;
    $self->{_x_axis_formula} = $name_formula;
    $self->{_x_axis_data_id} = $data_id;
}


###############################################################################
#
# set_y_axis()
#
# Set the properties of the Y-axis.
#
sub set_y_axis {

    my $self = shift;
    my %arg  = @_;

    my ( $name, $name_formula ) =
      $self->_process_names( $arg{name}, $arg{name_formula} );

    my $data_id = $self->_get_data_id( $name_formula, $arg{data} );

    $self->{_y_axis_name}    = $name;
    $self->{_y_axis_formula} = $name_formula;
    $self->{_y_axis_data_id} = $data_id;
}


###############################################################################
#
# set_title()
#
# Set the properties of the chart title.
#
sub set_title {

    my $self = shift;
    my %arg  = @_;

    my ( $name, $name_formula ) =
      $self->_process_names( $arg{name}, $arg{name_formula} );

    my $data_id = $self->_get_data_id( $name_formula, $arg{data} );

    $self->{_title_name}    = $name;
    $self->{_title_formula} = $name_formula;
    $self->{_title_data_id} = $data_id;
}


###############################################################################
#
# set_legend()
#
# Set the properties of the chart legend.
#
sub set_legend {

    my $self = shift;
    my %arg  = @_;

    $self->{_legend_position} = $arg{position} // 'right';
}


###############################################################################
#
# set_plotarea()
#
# Set the properties of the chart plotarea.
#
sub set_plotarea {

    # TODO. Need to refactor for XLSX format.

    my $self = shift;
    my %arg  = @_;
    return unless keys %arg;

    my $area = $self->{_plotarea};

    # Set the plotarea visibility.
    if ( defined $arg{visible} ) {
        $area->{_visible} = $arg{visible};
        return if !$area->{_visible};
    }

    # TODO. could move this out of if statement.
    $area->{_bg_color_index} = 0x08;

    # Set the chart background colour.
    if ( defined $arg{color} ) {
        my ( $index, $rgb ) = $self->_get_color_indices( $arg{color} );
        if ( defined $index ) {
            $area->{_fg_color_index} = $index;
            $area->{_fg_color_rgb}   = $rgb;
            $area->{_bg_color_index} = 0x08;
            $area->{_bg_color_rgb}   = 0x000000;
        }
    }

    # Set the border line colour.
    if ( defined $arg{line_color} ) {
        my ( $index, $rgb ) = $self->_get_color_indices( $arg{line_color} );
        if ( defined $index ) {
            $area->{_line_color_index} = $index;
            $area->{_line_color_rgb}   = $rgb;
        }
    }

    # Set the border line pattern.
    if ( defined $arg{line_pattern} ) {
        my $pattern = $self->_get_line_pattern( $arg{line_pattern} );
        $area->{_line_pattern} = $pattern;
    }

    # Set the border line weight.
    if ( defined $arg{line_weight} ) {
        my $weight = $self->_get_line_weight( $arg{line_weight} );
        $area->{_line_weight} = $weight;
    }
}


###############################################################################
#
# set_chartarea()
#
# Set the properties of the chart chartarea.
#
sub set_chartarea {

    # TODO. Need to refactor for XLSX format.

    my $self = shift;
    my %arg  = @_;
    return unless keys %arg;

    my $area = $self->{_chartarea};

    # Embedded automatic line weight has a different default value.
    $area->{_line_weight} = 0xFFFF if $self->{_embedded};


    # Set the chart background colour.
    if ( defined $arg{color} ) {
        my ( $index, $rgb ) = $self->_get_color_indices( $arg{color} );
        if ( defined $index ) {
            $area->{_fg_color_index} = $index;
            $area->{_fg_color_rgb}   = $rgb;
            $area->{_bg_color_index} = 0x08;
            $area->{_bg_color_rgb}   = 0x000000;
            $area->{_area_pattern}   = 1;
            $area->{_area_options}   = 0x0000 if $self->{_embedded};
            $area->{_visible}        = 1;
        }
    }

    # Set the border line colour.
    if ( defined $arg{line_color} ) {
        my ( $index, $rgb ) = $self->_get_color_indices( $arg{line_color} );
        if ( defined $index ) {
            $area->{_line_color_index} = $index;
            $area->{_line_color_rgb}   = $rgb;
            $area->{_line_pattern}     = 0x00;
            $area->{_line_options}     = 0x0000;
            $area->{_visible}          = 1;
        }
    }

    # Set the border line pattern.
    if ( defined $arg{line_pattern} ) {
        my $pattern = $self->_get_line_pattern( $arg{line_pattern} );
        $area->{_line_pattern}     = $pattern;
        $area->{_line_options}     = 0x0000;
        $area->{_line_color_index} = 0x4F if !defined $arg{line_color};
        $area->{_visible}          = 1;
    }

    # Set the border line weight.
    if ( defined $arg{line_weight} ) {
        my $weight = $self->_get_line_weight( $arg{line_weight} );
        $area->{_line_weight}      = $weight;
        $area->{_line_options}     = 0x0000;
        $area->{_line_pattern}     = 0x00 if !defined $arg{line_pattern};
        $area->{_line_color_index} = 0x4F if !defined $arg{line_color};
        $area->{_visible}          = 1;
    }
}


###############################################################################
#
# set_style()
#
# Set on of the 42 built-in Excel chart styles. The default style is 2.
#
sub set_style {

    my $self = shift;
    my $style_id = shift // 2;

    if ( $style_id < 0 || $style_id > 42 ) {
        $style_id = 2;
    }

    $self->{_style_id} = $style_id;
}


###############################################################################
#
# Internal methods. The following section of methods are used for the internal
# structuring of the Chart object and file format.
#
###############################################################################


###############################################################################
#
# _aref_to_formula()
#
# Convert and aref of row col values to a range formula.
#
sub _aref_to_formula {

    my $self = shift;
    my $data = shift;

    # If it isn't an array ref it is probably a formula already.
    return $data if !ref $data;

    my $formula = xl_range_formula( @$data );

    return $formula;
}


###############################################################################
#
# _process_names()
#
# Switch name and name_formula parameters if required.
#
sub _process_names {

    my $self         = shift;
    my $name         = shift;
    my $name_formula = shift;

    # Name looks like a formula, use it to set name_formula.
    if ( defined $name && $name =~ m/^=[^!]+!\$/ ) {
        $name_formula = $name;
        $name         = '';
    }

    return ( $name, $name_formula );
}


###############################################################################
#
# _get_data_type()
#
# Find the overall type of the data associated with a series.
#
# TODO. Need to handle date type.
#
sub _get_data_type {

    my $self = shift;
    my $data = shift;

    # Check for no data in the series.
    return 'none' if !defined $data;
    return 'none' if @$data == 0;

    # If the token isn't a number assume it is a string.
    for my $token ( @$data ) {
        return 'str'
          if $token !~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/;
    }

    # The series data was all numeric.
    return 'num';
}


###############################################################################
#
# _get_data_id()
#
# Assign an id to a each unique series formula or title/axis formula. Repeated
# formulas such as for categories get the same id. If the series or title
# has user specified data associated with it then that is also stored. This
# data is used to populate cached Excel data when creating a chart.
# If there is no user defined data then it will be populated by the parent
# workbook in Workbook::_add_chart_data()
#
sub _get_data_id {

    my $self    = shift;
    my $formula = shift;
    my $data    = shift;
    my $id;

    # Ignore series without a range formula.
    return unless $formula;

    # Strip the leading '=' from the formula.
    $formula =~ s/^=//;

    # Store the data id in a hash keyed by the formula and store the data
    # in a separate array with the same id.
    if ( !exists $self->{_formula_ids}->{$formula} ) {

        # Haven't seen this formula before.
        $id = @{ $self->{_formula_data} };

        push @{ $self->{_formula_data} }, $data;
        $self->{_formula_ids}->{$formula} = $id;
    }
    else {

        # Formula already seen. Return existing id.
        $id = $self->{_formula_ids}->{$formula};

        # Store user defined data if it isn't already there.
        if ( !defined $self->{_formula_data}->[$id] ) {
            $self->{_formula_data}->[$id] = $data;
        }
    }

    return $id;
}


###############################################################################
#
# _get_color_indices()
#
# Convert the user specified colour index or string to an colour index and
# RGB colour number.
#
sub _get_color_indices {

    my $self  = shift;
    my $color = shift;
    my $index;
    my $rgb;

    return ( undef, undef ) if !defined $color;

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


    # Check for the various supported colour index/name possibilities.
    if ( exists $colors{$color} ) {

        # Colour matches one of the supported colour names.
        $index = $colors{$color};
    }
    elsif ( $color =~ m/\D/ ) {

        # Return undef if $color is a string but not one of the supported ones.
        return ( undef, undef );
    }
    elsif ( $color < 8 || $color > 63 ) {

        # Return undef if index is out of range.
        return ( undef, undef );
    }
    else {

        # We should have a valid color index in a valid range.
        $index = $color;
    }

    $rgb = $self->_get_color_rbg( $index );
    return ( $index, $rgb );
}


###############################################################################
#
# _get_color_rbg()
#
# Get the RedGreenBlue number for the colour index from the Workbook palette.
#
sub _get_color_rbg {

    my $self  = shift;
    my $index = shift;

    # Adjust colour index from 8-63 (user range) to 0-55 (Excel range).
    $index -= 8;

    my @red_green_blue = @{ $self->{_palette}->[$index] };
    return unpack 'V', pack 'C*', @red_green_blue;
}


###############################################################################
#
# _get_line_pattern()
#
# Get the Excel chart index for line pattern that corresponds to the user
# defined value.
#
sub _get_line_pattern {

    my $self    = shift;
    my $value   = lc shift;
    my $default = 0;
    my $pattern;

    my %patterns = (
        0              => 5,
        1              => 0,
        2              => 1,
        3              => 2,
        4              => 3,
        5              => 4,
        6              => 7,
        7              => 6,
        8              => 8,
        'solid'        => 0,
        'dash'         => 1,
        'dot'          => 2,
        'dash-dot'     => 3,
        'dash-dot-dot' => 4,
        'none'         => 5,
        'dark-gray'    => 6,
        'medium-gray'  => 7,
        'light-gray'   => 8,
    );

    if ( exists $patterns{$value} ) {
        $pattern = $patterns{$value};
    }
    else {
        $pattern = $default;
    }

    return $pattern;
}


###############################################################################
#
# _get_line_weight()
#
# Get the Excel chart index for line weight that corresponds to the user
# defined value.
#
sub _get_line_weight {

    my $self    = shift;
    my $value   = lc shift;
    my $default = 0;
    my $weight;

    my %weights = (
        1          => -1,
        2          => 0,
        3          => 1,
        4          => 2,
        'hairline' => -1,
        'narrow'   => 0,
        'medium'   => 1,
        'wide'     => 2,
    );

    if ( exists $weights{$value} ) {
        $weight = $weights{$value};
    }
    else {
        $weight = $default;
    }

    return $weight;
}


###############################################################################
#
# _add_axis_id()
#
# Add a unique id for an axis.
#
sub _add_axis_id {

    my $self       = shift;
    my $chart_id   = 1 + $self->{_id};
    my $axis_count = 1 + @{ $self->{_axis_ids} };

    my $axis_id = sprintf '5%03d%04d', $chart_id, $axis_count;

    push @{ $self->{_axis_ids} }, $axis_id;

    return $axis_id;
}


###############################################################################
#
# Config data.
#
###############################################################################


###############################################################################
#
# _set_default_properties()
#
# Setup the default properties for a chart.
#
sub _set_default_properties {

    my $self = shift;

    $self->{_chartarea} = {
        _visible          => 0,
        _fg_color_index   => 0x4E,
        _fg_color_rgb     => 0xFFFFFF,
        _bg_color_index   => 0x4D,
        _bg_color_rgb     => 0x000000,
        _area_pattern     => 0x0000,
        _area_options     => 0x0000,
        _line_pattern     => 0x0005,
        _line_weight      => 0xFFFF,
        _line_color_index => 0x4D,
        _line_color_rgb   => 0x000000,
        _line_options     => 0x0008,
    };

    $self->{_plotarea} = {
        _visible          => 1,
        _fg_color_index   => 0x16,
        _fg_color_rgb     => 0xC0C0C0,
        _bg_color_index   => 0x4F,
        _bg_color_rgb     => 0x000000,
        _area_pattern     => 0x0001,
        _area_options     => 0x0000,
        _line_pattern     => 0x0000,
        _line_weight      => 0x0000,
        _line_color_index => 0x17,
        _line_color_rgb   => 0x808080,
        _line_options     => 0x0000,
    };
}


###############################################################################
#
# _set_embedded_config_data()
#
# Setup the default configuration data for an embedded chart.
#
sub _set_embedded_config_data {

    my $self = shift;

    $self->{_embedded} = 1;

    # TODO. We may be able to remove this after refactoring.

    $self->{_chartarea} = {
        _visible          => 1,
        _fg_color_index   => 0x4E,
        _fg_color_rgb     => 0xFFFFFF,
        _bg_color_index   => 0x4D,
        _bg_color_rgb     => 0x000000,
        _area_pattern     => 0x0001,
        _area_options     => 0x0001,
        _line_pattern     => 0x0000,
        _line_weight      => 0x0000,
        _line_color_index => 0x4D,
        _line_color_rgb   => 0x000000,
        _line_options     => 0x0009,
    };

}


###############################################################################
#
# XML writing methods.
#
###############################################################################


##############################################################################
#
# _write_chart_space()
#
# Write the <c:chartSpace> element.
#
sub _write_chart_space {

    my $self    = shift;
    my $schema  = 'http://schemas.openxmlformats.org/';
    my $xmlns_c = $schema . 'drawingml/2006/chart';
    my $xmlns_a = $schema . 'drawingml/2006/main';
    my $xmlns_r = $schema . 'officeDocument/2006/relationships';

    my @attributes = (
        'xmlns:c' => $xmlns_c,
        'xmlns:a' => $xmlns_a,
        'xmlns:r' => $xmlns_r,
    );

    $self->{_writer}->startTag( 'c:chartSpace', @attributes );
}


##############################################################################
#
# _write_lang()
#
# Write the <c:lang> element.
#
sub _write_lang {

    my $self = shift;
    my $val  = 'en-US';

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:lang', @attributes );
}


##############################################################################
#
# _write_style()
#
# Write the <c:style> element.
#
sub _write_style {

    my $self     = shift;
    my $style_id = $self->{_style_id};

    # Don't write an element for the default style, 2.
    return if $style_id == 2;

    my @attributes = ( 'val' => $style_id );

    $self->{_writer}->emptyTag( 'c:style', @attributes );
}


##############################################################################
#
# _write_chart()
#
# Write the <c:chart> element.
#
sub _write_chart {

    my $self = shift;

    $self->{_writer}->startTag( 'c:chart' );

    # Write the chart title elements.
    my $title;
    if ( $title = $self->{_title_formula} ) {
        $self->_write_title_formula( $title, $self->{_title_data_id} );
    }
    elsif ( $title = $self->{_title_name} ) {
        $self->_write_title_rich( $title );
    }

    # Write the c:plotArea element.
    $self->_write_plot_area();

    # Write the c:legend element.
    $self->_write_legend();

    # Write the c:plotVisOnly element.
    $self->_write_plot_vis_only();

    $self->{_writer}->endTag( 'c:chart' );
}


##############################################################################
#
# _write_plot_area()
#
# Write the <c:plotArea> element.
#
sub _write_plot_area {

    my $self = shift;

    $self->{_writer}->startTag( 'c:plotArea' );

    # Write the c:layout element.
    $self->_write_layout();

    # Write the subclass chart type element.
    $self->_write_chart_type();

    # Write the c:catAx element.
    $self->_write_cat_axis();

    # Write the c:catAx element.
    $self->_write_val_axis();

    $self->{_writer}->endTag( 'c:plotArea' );
}


##############################################################################
#
# _write_layout()
#
# Write the <c:layout> element.
#
sub _write_layout {

    my $self = shift;

    $self->{_writer}->emptyTag( 'c:layout' );
}


##############################################################################
#
# _write_chart_type()
#
# Write the chart type element. This method should be overridden by the
# subclasses.
#
sub _write_chart_type {

    my $self = shift;
}


##############################################################################
#
# _write_grouping()
#
# Write the <c:grouping> element.
#
sub _write_grouping {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:grouping', @attributes );
}


##############################################################################
#
# _write_series()
#
# Write the series elements.
#
sub _write_series {

    my $self = shift;

    # Write each series with subelements.
    my $index = 0;
    for my $series ( @{ $self->{_series} } ) {
        $self->_write_ser( $index++, $series );
    }

    # Write the c:marker element.
    $self->_write_marker_value();

    # Generate the axis ids.
    $self->_add_axis_id();
    $self->_add_axis_id();

    # Write the c:axId element.
    $self->_write_axis_id( $self->{_axis_ids}->[0] );
    $self->_write_axis_id( $self->{_axis_ids}->[1] );
}


##############################################################################
#
# _write_ser()
#
# Write the <c:ser> element.
#
sub _write_ser {

    my $self       = shift;
    my $index      = shift;
    my $series     = shift;

    $self->{_writer}->startTag( 'c:ser' );

    # Write the c:idx element.
    $self->_write_idx( $index );

    # Write the c:order element.
    $self->_write_order( $index );

    # Write the series name.
    $self->_write_series_name( $series );

    # Write the c:marker element.
    $self->_write_marker();

    # Write the c:cat element.
    $self->_write_cat( $series );

    # Write the c:val element.
    $self->_write_val( $series );

    $self->{_writer}->endTag( 'c:ser' );
}


##############################################################################
#
# _write_idx()
#
# Write the <c:idx> element.
#
sub _write_idx {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:idx', @attributes );
}


##############################################################################
#
# _write_order()
#
# Write the <c:order> element.
#
sub _write_order {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:order', @attributes );
}


##############################################################################
#
# _write_series_name()
#
# Write the series name.
#
sub _write_series_name {

    my $self   = shift;
    my $series = shift;

    my $name;
    if ( $name = $series->{_name_formula} ) {
        $self->_write_tx_formula( $name, $series->{_name_id} );
    }
    elsif ( $name = $series->{_name} ) {
        $self->_write_tx_value( $name );
    }

}


##############################################################################
#
# _write_cat()
#
# Write the <c:cat> element.
#
sub _write_cat {

    my $self    = shift;
    my $series  = shift;
    my $formula = $series->{_categories};
    my $data_id = $series->{_cat_data_id};
    my $data;

    if ( defined $data_id ) {
        $data = $self->{_formula_data}->[$data_id];
    }

    # Ignore <c:cat> elements for charts without category values.
    return unless $formula;

    $self->{_has_category} = 1;

    $self->{_writer}->startTag( 'c:cat' );

    # Check the type of cached data.
    my $type = $self->_get_data_type( $data );

    if ( $type eq 'str' ) {

        $self->{_has_category} = 0;

        # Write the c:numRef element.
        $self->_write_str_ref( $formula, $data, $type );
    }
    else {

        # Write the c:numRef element.
        $self->_write_num_ref( $formula, $data, $type );
    }

    $self->{_writer}->endTag( 'c:cat' );
}


##############################################################################
#
# _write_val()
#
# Write the <c:val> element.
#
sub _write_val {

    my $self    = shift;
    my $series  = shift;
    my $formula = $series->{_values};
    my $data_id = $series->{_val_data_id};
    my $data    = $self->{_formula_data}->[$data_id];

    $self->{_writer}->startTag( 'c:val' );

    # Check the type of cached data.
    my $type = $self->_get_data_type( $data );

    if ( $type eq 'str' ) {

        # Write the c:numRef element.
        $self->_write_str_ref( $formula, $data, $type );
    }
    else {

        # Write the c:numRef element.
        $self->_write_num_ref( $formula, $data, $type );
    }

    $self->{_writer}->endTag( 'c:val' );
}


##############################################################################
#
# _write_num_ref()
#
# Write the <c:numRef> element.
#
sub _write_num_ref {

    my $self    = shift;
    my $formula = shift;
    my $data    = shift;
    my $type    = shift;

    $self->{_writer}->startTag( 'c:numRef' );

    # Write the c:f element.
    $self->_write_series_formula( $formula );

    if ( $type eq 'num' ) {

        # Write the c:numCache element.
        $self->_write_num_cache( $data );
    }
    elsif ( $type eq 'str' ) {

        # Write the c:strCache element.
        $self->_write_str_cache( $data );
    }

    $self->{_writer}->endTag( 'c:numRef' );
}



##############################################################################
#
# _write_str_ref()
#
# Write the <c:strRef> element.
#
sub _write_str_ref {

    my $self    = shift;
    my $formula = shift;
    my $data    = shift;
    my $type    = shift;

    $self->{_writer}->startTag( 'c:strRef' );

    # Write the c:f element.
    $self->_write_series_formula( $formula );

    if ( $type eq 'num' ) {

        # Write the c:numCache element.
        $self->_write_num_cache( $data );
    }
    elsif ( $type eq 'str' ) {

        # Write the c:strCache element.
        $self->_write_str_cache( $data );
    }

    $self->{_writer}->endTag( 'c:strRef' );
}


##############################################################################
#
# _write_series_formula()
#
# Write the <c:f> element.
#
sub _write_series_formula {

    my $self    = shift;
    my $formula = shift;

    # Strip the leading '=' from the formula.
    $formula =~ s/^=//;

    $self->{_writer}->dataElement( 'c:f', $formula );
}


##############################################################################
#
# _write_axis_id()
#
# Write the <c:axId> element.
#
sub _write_axis_id {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:axId', @attributes );
}


##############################################################################
#
# _write_cat_axis()
#
# Write the <c:catAx> element.
#
sub _write_cat_axis {

    my $self     = shift;
    my $position = shift // $self->{_cat_axis_position};
    my $horiz    = $self->{_horiz_cat_axis};

    $self->{_writer}->startTag( 'c:catAx' );

    $self->_write_axis_id( $self->{_axis_ids}->[0] );

    # Write the c:scaling element.
    $self->_write_scaling();

    # Write the c:axPos element.
    $self->_write_axis_pos( $position );

    # Write the axis title elements.
    my $title;
    if ( $title = $self->{_x_axis_formula} ) {
        $self->_write_title_formula( $title, $self->{_x_axis_data_id}, $horiz );
    }
    elsif ( $title = $self->{_x_axis_name} ) {
        $self->_write_title_rich( $title, $horiz );
    }

    # Write the c:numFmt element.
    $self->_write_num_fmt();

    # Write the c:tickLblPos element.
    $self->_write_tick_label_pos( 'nextTo' );

    # Write the c:crossAx element.
    $self->_write_cross_axis( $self->{_axis_ids}->[1] );

    # Write the c:crosses element.
    $self->_write_crosses( 'autoZero' );

    # Write the c:auto element.
    $self->_write_auto( 1 );

    # Write the c:labelAlign element.
    $self->_write_label_align( 'ctr' );

    # Write the c:labelOffset element.
    $self->_write_label_offset( 100 );

    $self->{_writer}->endTag( 'c:catAx' );
}


##############################################################################
#
# _write_val_axis()
#
# Write the <c:valAx> element.
#
# TODO. Maybe should have a _write_cat_val_axis() method as well for scatter.
#
sub _write_val_axis {

    my $self                 = shift;
    my $position             = shift // $self->{_val_axis_position};
    my $hide_major_gridlines = shift;
    my $horiz                = $self->{_horiz_val_axis};

    $self->{_writer}->startTag( 'c:valAx' );

    $self->_write_axis_id( $self->{_axis_ids}->[1] );

    # Write the c:scaling element.
    $self->_write_scaling();

    # Write the c:axPos element.
    $self->_write_axis_pos( $position );

    # Write the c:majorGridlines element.
    $self->_write_major_gridlines() if not $hide_major_gridlines;

    # Write the axis title elements.
    my $title;
    if ( $title = $self->{_y_axis_formula} ) {
        $self->_write_title_formula( $title, $self->{_y_axis_data_id}, $horiz );
    }
    elsif ( $title = $self->{_y_axis_name} ) {
        $self->_write_title_rich( $title, $horiz );
    }

    # Write the c:numberFormat element.
    $self->_write_number_format();

    # Write the c:tickLblPos element.
    $self->_write_tick_label_pos( 'nextTo' );

    # Write the c:crossAx element.
    $self->_write_cross_axis( $self->{_axis_ids}->[0] );

    # Write the c:crosses element.
    $self->_write_crosses( 'autoZero' );

    # Write the c:crossBetween element.
    $self->_write_cross_between();

    $self->{_writer}->endTag( 'c:valAx' );
}


##############################################################################
#
# _write_cat_val_axis()
#
# Write the <c:valAx> element. This is for the second valAx in scatter plots.
#
#
sub _write_cat_val_axis {

    my $self                 = shift;
    my $position             = shift // $self->{_val_axis_position};
    my $hide_major_gridlines = shift;
    my $horiz                = $self->{_horiz_val_axis};

    $self->{_writer}->startTag( 'c:valAx' );

    $self->_write_axis_id( $self->{_axis_ids}->[0] );

    # Write the c:scaling element.
    $self->_write_scaling();

    # Write the c:axPos element.
    $self->_write_axis_pos( $position );

    # Write the c:majorGridlines element.
    $self->_write_major_gridlines() if not $hide_major_gridlines;

    # Write the axis title elements.
    my $title;
    if ( $title = $self->{_x_axis_formula} ) {
        $self->_write_title_formula( $title, $self->{_y_axis_data_id}, $horiz );
    }
    elsif ( $title = $self->{_x_axis_name} ) {
        $self->_write_title_rich( $title, $horiz );
    }

    # Write the c:numberFormat element.
    $self->_write_number_format();

    # Write the c:tickLblPos element.
    $self->_write_tick_label_pos( 'nextTo' );

    # Write the c:crossAx element.
    $self->_write_cross_axis( $self->{_axis_ids}->[1] );

    # Write the c:crosses element.
    $self->_write_crosses( 'autoZero' );

    # Write the c:crossBetween element.
    $self->_write_cross_between();

    $self->{_writer}->endTag( 'c:valAx' );
}


##############################################################################
#
# _write_date_axis()
#
# Write the <c:dateAx> element.
#
sub _write_date_axis {

    my $self = shift;
    my $position = shift // $self->{_cat_axis_position};

    $self->{_writer}->startTag( 'c:dateAx' );

    $self->_write_axis_id( $self->{_axis_ids}->[0] );

    # Write the c:scaling element.
    $self->_write_scaling();

    # Write the c:axPos element.
    $self->_write_axis_pos( $position );

    # Write the axis title elements.
    my $title;
    if ( $title = $self->{_x_axis_formula} ) {
        $self->_write_title_formula( $title, $self->{_x_axis_data_id} );
    }
    elsif ( $title = $self->{_x_axis_name} ) {
        $self->_write_title_rich( $title );
    }

    # Write the c:numFmt element.
    $self->_write_num_fmt( 'dd/mm/yyyy' );

    # Write the c:tickLblPos element.
    $self->_write_tick_label_pos( 'nextTo' );

    # Write the c:crossAx element.
    $self->_write_cross_axis( $self->{_axis_ids}->[1] );

    # Write the c:crosses element.
    $self->_write_crosses( 'autoZero' );

    # Write the c:auto element.
    $self->_write_auto( 1 );

    # Write the c:labelOffset element.
    $self->_write_label_offset( 100 );

    $self->{_writer}->endTag( 'c:dateAx' );
}


##############################################################################
#
# _write_scaling()
#
# Write the <c:scaling> element.
#
sub _write_scaling {

    my $self = shift;

    $self->{_writer}->startTag( 'c:scaling' );

    # Write the c:orientation element.
    $self->_write_orientation();

    $self->{_writer}->endTag( 'c:scaling' );
}


##############################################################################
#
# _write_orientation()
#
# Write the <c:orientation> element.
#
sub _write_orientation {

    my $self = shift;
    my $val  = 'minMax';

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:orientation', @attributes );
}


##############################################################################
#
# _write_axis_pos()
#
# Write the <c:axPos> element.
#
sub _write_axis_pos {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:axPos', @attributes );
}


##############################################################################
#
# _write_num_fmt()
#
# Write the <c:numFmt> element.
#
sub _write_num_fmt {

    my $self          = shift;
    my $format_code   = shift // 'General';
    my $source_linked = 1;

    # These elements are only required for charts with categories.
    return unless $self->{_has_category};

    my @attributes = (
        'formatCode'   => $format_code,
        'sourceLinked' => $source_linked,
    );

    $self->{_writer}->emptyTag( 'c:numFmt', @attributes );
}


##############################################################################
#
# _write_tick_label_pos()
#
# Write the <c:tickLblPos> element.
#
sub _write_tick_label_pos {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:tickLblPos', @attributes );
}


##############################################################################
#
# _write_cross_axis()
#
# Write the <c:crossAx> element.
#
sub _write_cross_axis {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:crossAx', @attributes );
}


##############################################################################
#
# _write_crosses()
#
# Write the <c:crosses> element.
#
sub _write_crosses {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:crosses', @attributes );
}


##############################################################################
#
# _write_auto()
#
# Write the <c:auto> element.
#
sub _write_auto {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:auto', @attributes );
}


##############################################################################
#
# _write_label_align()
#
# Write the <c:labelAlign> element.
#
sub _write_label_align {

    my $self = shift;
    my $val  = 'ctr';

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:lblAlgn', @attributes );
}


##############################################################################
#
# _write_label_offset()
#
# Write the <c:labelOffset> element.
#
sub _write_label_offset {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:lblOffset', @attributes );
}


##############################################################################
#
# _write_major_gridlines()
#
# Write the <c:majorGridlines> element.
#
sub _write_major_gridlines {

    my $self = shift;

    $self->{_writer}->emptyTag( 'c:majorGridlines' );
}


##############################################################################
#
# _write_number_format()
#
# Write the <c:numberFormat> element.
#
# TODO. Merge/replace with _write_num_fmt().
#
sub _write_number_format {

    my $self          = shift;
    my $format_code   = 'General';
    my $source_linked = 1;

    my @attributes = (
        'formatCode'   => $format_code,
        'sourceLinked' => $source_linked,
    );

    $self->{_writer}->emptyTag( 'c:numFmt', @attributes );
}

##############################################################################
#
# _write_cross_between()
#
# Write the <c:crossBetween> element.
#
sub _write_cross_between {

    my $self = shift;
    my $val  = $self->{_cross_between} // 'between';

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:crossBetween', @attributes );
}


##############################################################################
#
# _write_legend()
#
# Write the <c:legend> element.
#
sub _write_legend {

    my $self = shift;
    my $position = $self->{_legend_position};
    my $overlay = 0;

    if ($position =~ s/^overlay_//) {
        $overlay = 1;
    }

    my %allowed = (
        right  => 'r',
        left   => 'l',
        top    => 't',
        bottom => 'b',
    );

    return if $position eq 'none';
    return unless exists $allowed{$position};

    $position = $allowed{$position};

    $self->{_writer}->startTag( 'c:legend' );

    # Write the c:legendPos element.
    $self->_write_legend_pos( $position );

    # Write the c:layout element.
    $self->_write_layout();

    # Write the c:overlay element.
    $self->_write_overlay() if $overlay;

    $self->{_writer}->endTag( 'c:legend' );
}


##############################################################################
#
# _write_legend_pos()
#
# Write the <c:legendPos> element.
#
sub _write_legend_pos {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:legendPos', @attributes );
}


##############################################################################
#
# _write_overlay()
#
# Write the <c:overlay> element.
#
sub _write_overlay {

    my $self = shift;
    my $val  = 1;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:overlay', @attributes );
}


##############################################################################
#
# _write_plot_vis_only()
#
# Write the <c:plotVisOnly> element.
#
sub _write_plot_vis_only {

    my $self = shift;
    my $val  = 1;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:plotVisOnly', @attributes );
}


##############################################################################
#
# _write_print_settings()
#
# Write the <c:printSettings> element.
#
sub _write_print_settings {

    my $self = shift;

    $self->{_writer}->startTag( 'c:printSettings' );

    # Write the c:headerFooter element.
    $self->_write_header_footer();

    # Write the c:pageMargins element.
    $self->_write_page_margins();

    # Write the c:pageSetup element.
    $self->_write_page_setup();

    $self->{_writer}->endTag( 'c:printSettings' );
}


##############################################################################
#
# _write_header_footer()
#
# Write the <c:headerFooter> element.
#
sub _write_header_footer {

    my $self = shift;

    $self->{_writer}->emptyTag( 'c:headerFooter' );
}


##############################################################################
#
# _write_page_margins()
#
# Write the <c:pageMargins> element.
#
sub _write_page_margins {

    my $self   = shift;
    my $b      = 0.75;
    my $l      = 0.7;
    my $r      = 0.7;
    my $t      = 0.75;
    my $header = 0.3;
    my $footer = 0.3;

    my @attributes = (
        'b'      => $b,
        'l'      => $l,
        'r'      => $r,
        't'      => $t,
        'header' => $header,
        'footer' => $footer,
    );

    $self->{_writer}->emptyTag( 'c:pageMargins', @attributes );
}


##############################################################################
#
# _write_page_setup()
#
# Write the <c:pageSetup> element.
#
sub _write_page_setup {

    my $self = shift;

    $self->{_writer}->emptyTag( 'c:pageSetup' );
}


##############################################################################
#
# _write_title_rich()
#
# Write the <c:title> element for a rich string.
#
sub _write_title_rich {

    my $self  = shift;
    my $title = shift;
    my $horiz = shift;

    $self->{_writer}->startTag( 'c:title' );

    # Write the c:tx element.
    $self->_write_tx_rich( $title, $horiz );

    # Write the c:layout element.
    $self->_write_layout();

    $self->{_writer}->endTag( 'c:title' );
}


##############################################################################
#
# _write_title_formula()
#
# Write the <c:title> element for a rich string.
#
sub _write_title_formula {

    my $self    = shift;
    my $title   = shift;
    my $data_id = shift;
    my $horiz   = shift;

    $self->{_writer}->startTag( 'c:title' );

    # Write the c:tx element.
    $self->_write_tx_formula( $title, $data_id );

    # Write the c:layout element.
    $self->_write_layout();

    # Write the c:txPr element.
    $self->_write_tx_pr( $horiz );

    $self->{_writer}->endTag( 'c:title' );
}


##############################################################################
#
# _write_tx_rich()
#
# Write the <c:tx> element.
#
sub _write_tx_rich {

    my $self  = shift;
    my $title = shift;
    my $horiz = shift;

    $self->{_writer}->startTag( 'c:tx' );

    # Write the c:rich element.
    $self->_write_rich( $title, $horiz );

    $self->{_writer}->endTag( 'c:tx' );
}



##############################################################################
#
# _write_tx_value()
#
# Write the <c:tx> element with a simple value such as for series names.
#
sub _write_tx_value {

    my $self  = shift;
    my $title = shift;

    $self->{_writer}->startTag( 'c:tx' );

    # Write the c:v element.
    $self->_write_v( $title );

    $self->{_writer}->endTag( 'c:tx' );
}


##############################################################################
#
# _write_tx_formula()
#
# Write the <c:tx> element.
#
sub _write_tx_formula {

    my $self    = shift;
    my $title   = shift;
    my $data_id = shift;
    my $data;

    if ( defined $data_id ) {
        $data = $self->{_formula_data}->[$data_id];
    }

    $self->{_writer}->startTag( 'c:tx' );

    # Write the c:strRef element.
    $self->_write_str_ref( $title, $data, 'str' );

    $self->{_writer}->endTag( 'c:tx' );
}


##############################################################################
#
# _write_rich()
#
# Write the <c:rich> element.
#
sub _write_rich {

    my $self  = shift;
    my $title = shift;
    my $horiz = shift;

    $self->{_writer}->startTag( 'c:rich' );

    # Write the a:bodyPr element.
    $self->_write_a_body_pr( $horiz );

    # Write the a:lstStyle element.
    $self->_write_a_lst_style();

    # Write the a:p element.
    $self->_write_a_p_rich( $title );


    $self->{_writer}->endTag( 'c:rich' );
}


##############################################################################
#
# _write_a_body_pr()
#
# Write the <a:bodyPr> element.
#
sub _write_a_body_pr {

    my $self  = shift;
    my $horiz = shift;
    my $rot   = -5400000;
    my $vert  = 'horz';

    my @attributes = (
        'rot'  => $rot,
        'vert' => $vert,
    );

    @attributes = () if !$horiz;

    $self->{_writer}->emptyTag( 'a:bodyPr', @attributes );
}


##############################################################################
#
# _write_a_lst_style()
#
# Write the <a:lstStyle> element.
#
sub _write_a_lst_style {

    my $self = shift;

    $self->{_writer}->emptyTag( 'a:lstStyle' );
}


##############################################################################
#
# _write_a_p_rich()
#
# Write the <a:p> element for rich string titles.
#
sub _write_a_p_rich {

    my $self  = shift;
    my $title = shift;

    $self->{_writer}->startTag( 'a:p' );

    # Write the a:pPr element.
    $self->_write_a_p_pr_rich();

    # Write the a:r element.
    $self->_write_a_r( $title );

    $self->{_writer}->endTag( 'a:p' );
}


##############################################################################
#
# _write_a_p_formula()
#
# Write the <a:p> element for formula titles.
#
sub _write_a_p_formula {

    my $self  = shift;
    my $title = shift;

    $self->{_writer}->startTag( 'a:p' );

    # Write the a:pPr element.
    $self->_write_a_p_pr_formula();

    # Write the a:endParaRPr element.
    $self->_write_a_end_para_rpr();

    $self->{_writer}->endTag( 'a:p' );
}


##############################################################################
#
# _write_a_p_pr_rich()
#
# Write the <a:pPr> element for rich string titles.
#
sub _write_a_p_pr_rich {

    my $self = shift;

    $self->{_writer}->startTag( 'a:pPr' );

    # Write the a:defRPr element.
    $self->_write_a_def_rpr();

    $self->{_writer}->endTag( 'a:pPr' );
}


##############################################################################
#
# _write_a_p_pr_formula()
#
# Write the <a:pPr> element for formula titles.
#
sub _write_a_p_pr_formula {

    my $self = shift;

    $self->{_writer}->startTag( 'a:pPr' );

    # Write the a:defRPr element.
    $self->_write_a_def_rpr();

    $self->{_writer}->endTag( 'a:pPr' );
}


##############################################################################
#
# _write_a_def_rpr()
#
# Write the <a:defRPr> element.
#
sub _write_a_def_rpr {

    my $self = shift;

    $self->{_writer}->emptyTag( 'a:defRPr' );
}


##############################################################################
#
# _write_a_end_para_rpr()
#
# Write the <a:endParaRPr> element.
#
sub _write_a_end_para_rpr {

    my $self = shift;
    my $lang = 'en-US';

    my @attributes = ( 'lang' => $lang );

    $self->{_writer}->emptyTag( 'a:endParaRPr', @attributes );
}


##############################################################################
#
# _write_a_r()
#
# Write the <a:r> element.
#
sub _write_a_r {

    my $self  = shift;
    my $title = shift;

    $self->{_writer}->startTag( 'a:r' );

    # Write the a:rPr element.
    $self->_write_a_r_pr();

    # Write the a:t element.
    $self->_write_a_t( $title );

    $self->{_writer}->endTag( 'a:r' );
}


##############################################################################
#
# _write_a_r_pr()
#
# Write the <a:rPr> element.
#
sub _write_a_r_pr {

    my $self = shift;
    my $lang = 'en-US';

    my @attributes = ( 'lang' => $lang, );

    $self->{_writer}->emptyTag( 'a:rPr', @attributes );
}


##############################################################################
#
# _write_a_t()
#
# Write the <a:t> element.
#
sub _write_a_t {

    my $self  = shift;
    my $title = shift;

    $self->{_writer}->dataElement( 'a:t', $title );
}


##############################################################################
#
# _write_tx_pr()
#
# Write the <c:txPr> element.
#
sub _write_tx_pr {

    my $self  = shift;
    my $horiz = shift;

    $self->{_writer}->startTag( 'c:txPr' );

    # Write the a:bodyPr element.
    $self->_write_a_body_pr( $horiz );

    # Write the a:lstStyle element.
    $self->_write_a_lst_style();

    # Write the a:p element.
    $self->_write_a_p_formula();

    $self->{_writer}->endTag( 'c:txPr' );
}


##############################################################################
#
# _write_marker()
#
# Write the <c:marker> element.
#
sub _write_marker {

    my $self  = shift;
    my $style = shift // $self->{_default_marker};

    return unless $style;

    $self->{_writer}->startTag( 'c:marker' );

    # Write the c:symbol element.
    $self->_write_symbol( $style );

    # Write the c:size element.
    $self->_write_marker_size( 3 ) if $style eq 'dot';

    $self->{_writer}->endTag( 'c:marker' );
}


##############################################################################
#
# _write_marker_value()
#
# Write the <c:marker> element without a sub-element.
#
sub _write_marker_value {

    my $self  = shift;
    my $style = $self->{_default_marker};

    return unless $style;

    my @attributes = ( 'val' => 1 );

    $self->{_writer}->emptyTag( 'c:marker', @attributes );
}


##############################################################################
#
# _write_marker_size()
#
# Write the <c:size> element.
#
sub _write_marker_size {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:size', @attributes );
}


##############################################################################
#
# _write_symbol()
#
# Write the <c:symbol> element.
#
sub _write_symbol {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:symbol', @attributes );
}


##############################################################################
#
# _write_sp_pr()
#
# Write the <c:spPr> element.
#
sub _write_sp_pr {

    my $self = shift;

    $self->{_writer}->startTag( 'c:spPr' );

    # Write the a:ln element.
    $self->_write_a_ln();

    $self->{_writer}->endTag( 'c:spPr' );
}


##############################################################################
#
# _write_a_ln()
#
# Write the <a:ln> element.
#
sub _write_a_ln {

    my $self = shift;
    my $w    = 28575;

    my @attributes = ( 'w' => $w );

    $self->{_writer}->startTag( 'a:ln', @attributes );

    # Write the a:noFill element.
    $self->_write_a_no_fill();

    $self->{_writer}->endTag( 'a:ln' );
}


##############################################################################
#
# _write_a_no_fill()
#
# Write the <a:noFill> element.
#
sub _write_a_no_fill {

    my $self = shift;

    $self->{_writer}->emptyTag( 'a:noFill' );
}


##############################################################################
#
# _write_hi_low_lines()
#
# Write the <c:hiLowLines> element.
#
sub _write_hi_low_lines {

    my $self = shift;

    $self->{_writer}->emptyTag( 'c:hiLowLines' );
}


##############################################################################
#
# _write_overlap()
#
# Write the <c:overlap> element.
#
sub _write_overlap {

    my $self = shift;
    my $val  = 100;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:overlap', @attributes );
}


##############################################################################
#
# _write_num_cache()
#
# Write the <c:numCache> element.
#
sub _write_num_cache {

    my $self  = shift;
    my $data  = shift;
    my $count = @$data;

    $self->{_writer}->startTag( 'c:numCache' );

    # Write the c:formatCode element.
    $self->_write_format_code( 'General' );

    # Write the c:ptCount element.
    $self->_write_pt_count( $count );

    for my $i ( 0 .. $count - 1 ) {

        # Write the c:pt element.
        $self->_write_pt( $i, $data->[$i] );
    }

    $self->{_writer}->endTag( 'c:numCache' );
}


##############################################################################
#
# _write_str_cache()
#
# Write the <c:strCache> element.
#
sub _write_str_cache {

    my $self  = shift;
    my $data  = shift;
    my $count = @$data;

    $self->{_writer}->startTag( 'c:strCache' );

    # Write the c:ptCount element.
    $self->_write_pt_count( $count );

    for my $i ( 0 .. $count - 1 ) {

        # Write the c:pt element.
        $self->_write_pt( $i, $data->[$i] );
    }

    $self->{_writer}->endTag( 'c:strCache' );
}


##############################################################################
#
# _write_format_code()
#
# Write the <c:formatCode> element.
#
sub _write_format_code {

    my $self = shift;
    my $data = shift;

    $self->{_writer}->dataElement( 'c:formatCode', $data );
}


##############################################################################
#
# _write_pt_count()
#
# Write the <c:ptCount> element.
#
sub _write_pt_count {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:ptCount', @attributes );
}


##############################################################################
#
# _write_pt()
#
# Write the <c:pt> element.
#
sub _write_pt {

    my $self  = shift;
    my $idx   = shift;
    my $value = shift;

    my @attributes = ( 'idx' => $idx );

    $self->{_writer}->startTag( 'c:pt', @attributes );

    # Write the c:v element.
    $self->_write_v( $value );

    $self->{_writer}->endTag( 'c:pt' );
}


##############################################################################
#
# _write_v()
#
# Write the <c:v> element.
#
sub _write_v {

    my $self = shift;
    my $data = shift;

    $self->{_writer}->dataElement( 'c:v', $data );
}


##############################################################################
#
# _write_protection()
#
# Write the <c:protection> element.
#
sub _write_protection {

    my $self = shift;

    return unless $self->{_protection};

    $self->{_writer}->emptyTag( 'c:protection' );
}


1;

__END__


=head1 NAME

Chart - A class for writing Excel Charts.

=head1 SYNOPSIS

To create a simple Excel file with a chart using Excel::Writer::XLSX:

    #!/usr/bin/perl

    use strict;
    use warnings;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'chart.xlsx' );
    my $worksheet = $workbook->add_worksheet();

    # Add the worksheet data the chart refers to.
    my $data = [
        [ 'Category', 2, 3, 4, 5, 6, 7 ],
        [ 'Value',    1, 4, 5, 2, 1, 5 ],

    ];

    $worksheet->write( 'A1', $data );

    # Add a worksheet chart.
    my $chart = $workbook->add_chart( type => 'column' );

    # Configure the chart.
    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$7',
        values     => '=Sheet1!$B$2:$B$7',
    );

    __END__


=head1 DESCRIPTION

The C<Chart> module is an abstract base class for modules that implement charts in L<Excel::Writer::XLSX>. The information below is applicable to all of the available subclasses.

The C<Chart> module isn't used directly, a chart object is created via the Workbook C<add_chart()> method where the chart type is specified:

    my $chart = $workbook->add_chart( type => 'column' );

Currently the supported chart types are:

=over

=item * C<area>

Creates an Area (filled line) style chart. See L<Excel::Writer::XLSX::Chart::Area>.

=item * C<bar>

Creates a Bar style (transposed histogram) chart. See L<Excel::Writer::XLSX::Chart::Bar>.

=item * C<column>

Creates a column style (histogram) chart. See L<Excel::Writer::XLSX::Chart::Column>.

=item * C<line>

Creates a Line style chart. See L<Excel::Writer::XLSX::Chart::Line>.

=item * C<pie>

Creates an Pie style chart. See L<Excel::Writer::XLSX::Chart::Pie>.

=item * C<scatter>

Creates an Scatter style chart. See L<Excel::Writer::XLSX::Chart::Scatter>.

=item * C<stock>

Creates an Stock style chart. See L<Excel::Writer::XLSX::Chart::Stock>.

=item * C<...>

More charts and sub-types will be supported in time. See the L</TODO> section.

=back


=head1 CHART METHODS

Methods that are common to all chart types are documented below. See the documentation for each sub class for chart specific information.

=head2 add_series()

In an Excel chart a "series" is a collection of information such as values, x-axis labels and the name that define which data is plotted.

With a Excel::Writer::XLSX chart object the C<add_series()> method is used to set the properties for a series:

    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$10', # Optional, depending on chart type.
        values     => '=Sheet1!$B$2:$B$10', # Required for all chart types.
        name       => 'Series name',        # Optional.
    );

The properties that can be set are:

=over

=item * C<values>

This is the most important property of a series and must be set for every chart object. It links the chart with the worksheet data that it displays. A formula or array ref can be used for the data range, see below.

=item * C<categories>

This sets the chart category labels. The category is more or less the same as the X-axis. In most chart types the C<categories> property is optional and the chart will just assume a sequential series from C<1 .. n>.

=item * C<name>

Set the name for the series. The name is displayed in the chart legend and in the formula bar. The name property is optional and if it isn't supplied will default to C<Series 1 .. n>.

=back

The C<categories> and C<values> can take either a range formula such as C<=Sheet1!$A$2:$A$7> or, more usefully when generating the range programmatically, an array ref with zero indexed row/column values:

     [ $sheetname, $row_start, $row_end, $col_start, $col_end ]

The following are equivalent:

    $chart->add_series( categories => '=Sheet1!$A$2:$A$7'      ); # Same as ...
    $chart->add_series( categories => [ 'Sheet1', 1, 6, 0, 0 ] ); # Zero-indexed.

You can add more than one series to a chart, in fact some chart types such as C<stock> require it. The series numbering and order in the final chart is the same as the order in which that are added.

    # Add the first series.
    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$7',
        values     => '=Sheet1!$B$2:$B$7',
        name       => 'Test data series 1',
    );

    # Add another series. Category is the same but values are different.
    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$7',
        values     => '=Sheet1!$C$2:$C$7',
        name       => 'Test data series 2',
    );



=head2 set_x_axis()

The C<set_x_axis()> method is used to set properties of the X axis.

    $chart->set_x_axis( name => 'Sample length (m)' );

The properties that can be set are:

=over

=item * C<name>

Set the name (title or caption) for the axis. The name is displayed below the X axis. The name can also be a formula such as C<=Sheet1!$A$1>. The name property is optional. The default is to have no axis name.

=back

Additional axis properties such as range, divisions and ticks will be made available in later releases.


=head2 set_y_axis()

The C<set_y_axis()> method is used to set properties of the Y axis.

    $chart->set_y_axis( name => 'Sample weight (kg)' );

The properties that can be set are:

=over

=item * C<name>

Set the name (title or caption) for the axis. The name is displayed to the left of the Y axis. The name can also be a formula such as C<=Sheet1!$A$1>. The name property is optional. The default is to have no axis name.

=back

Additional axis properties such as range, divisions and ticks will be made available in later releases.

=head2 set_title()

The C<set_title()> method is used to set properties of the chart title.

    $chart->set_title( name => 'Year End Results' );

The properties that can be set are:

=over

=item * C<name>

Set the name (title) for the chart. The name is displayed above the chart. The name can also be a formula such as C<=Sheet1!$A$1>. The name property is optional. The default is to have no chart title.

=back


=head2 set_legend()

The C<set_legend()> method is used to set properties of the chart legend.

    $chart->set_legend( position => 'none' );

The properties that can be set are:

=over

=item * C<position>

Set the position of the chart legend.

    $chart->set_legend( position => 'bottom' );

The default legend position is C<right>. The available positions are:

    right
    left
    top
    bottom
    none
    overlay_right
    overlay_left

=back


=head2 set_chartarea()

The C<set_chartarea()> method is used to set the properties of the chart area.

This method isn't implemented yet and is only available in L<Spreadsheet::WriteExcel>. However, it can be simulated using the C<set_style()> method, see below.

=head2 set_plotarea()

The C<set_plotarea()> method is used to set properties of the plot area of a chart.

This method isn't implemented yet and is only available in L<Spreadsheet::WriteExcel>. However, it can be simulated using the C<set_style()> method, see below.

=head2 set_style()

The C<set_style()> method is used to set the style of the chart to one of the 42 built-in styles available on the 'Design' tab in Excel:

    $chart->set_style( 4 );

The default style is 2.


=head1 WORKSHEET METHODS

In Excel a chartsheet (i.e, a chart that isn't embedded) shares properties with data worksheets such as tab selection, headers, footers, margins and print properties.

In Excel::Writer::XLSX you can set chartsheet properties using the same methods that are used for Worksheet objects.

The following Worksheet methods are also available through a non-embedded Chart object:

    get_name()
    activate()
    select()
    hide()
    set_first_sheet()
    protect()
    set_zoom()
    set_tab_color()

    set_landscape()
    set_portrait()
    set_paper()
    set_margins()
    set_header()
    set_footer()

See L<Excel::Writer::XLSX> for a detailed explanation of these methods.

=head1 EXAMPLE

Here is a complete example that demonstrates some of the available features when creating a chart.

    #!/usr/bin/perl

    use strict;
    use warnings;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'chart.xlsx' );
    my $worksheet = $workbook->add_worksheet();
    my $bold      = $workbook->add_format( bold => 1 );

    # Add the worksheet data that the charts will refer to.
    my $headings = [ 'Number', 'Batch 1', 'Batch 2' ];
    my $data = [
        [ 2,  3,  4,  5,  6,  7 ],
        [ 10, 40, 50, 20, 10, 50 ],
        [ 30, 60, 70, 50, 40, 30 ],

    ];

    $worksheet->write( 'A1', $headings, $bold );
    $worksheet->write( 'A2', $data );

    # Create a new chart object. In this case an embedded chart.
    my $chart = $workbook->add_chart( type => 'column', embedded => 1 );

    # Configure the first series.
    $chart->add_series(
        name       => '=Sheet1!$B$1',
        categories => '=Sheet1!$A$2:$A$7',
        values     => '=Sheet1!$B$2:$B$7',
    );

    # Configure second series. Note alternative use of array ref to define
    # ranges: [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
    $chart->add_series(
        name       => '=Sheet1!$C$1',
        categories => [ 'Sheet1', 1, 6, 0, 0 ],
        values     => [ 'Sheet1', 1, 6, 2, 2 ],
    );

    # Add a chart title and some axis labels.
    $chart->set_title ( name => 'Results of sample analysis' );
    $chart->set_x_axis( name => 'Test number' );
    $chart->set_y_axis( name => 'Sample length (mm)' );

    # Set an Excel chart style. Blue colors with white outline and shadow.
    $chart->set_style( 11 );

    # Insert the chart into the worksheet (with an offset).
    $worksheet->insert_chart( 'D2', $chart, 25, 10 );

    __END__

=begin html

<p>This will produce a chart that looks like this:</p>

<p><center><img src="http://homepage.eircom.net/~jmcnamara/perl/images/2007/area1.jpg" width="527" height="320" alt="Chart example." /></center></p>

=end html


=head1 TODO

Charts in Excel::Writer::XLSX is under active development. More chart types and features will be added in time.

Features that are on the TODO list and will be added are:

=over

=item * Chart sub-types such as stacked and percent stacked.

=item * Colours and formatting options. For now try the C<set_style()> method.

=item * Axis controls, range limits, gridlines.

=item * 3D charts.

=item * Additional chart types such as Bubble and Radar.

=back

If you are interested in sponsoring a feature to have it implemented or expedited let me know.


=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MM-MMXI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

