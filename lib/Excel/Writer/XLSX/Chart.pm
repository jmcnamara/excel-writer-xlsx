package Excel::Writer::XLSX::Chart;

###############################################################################
#
# Chart - A writer class for Excel Charts.
#
#
# Used in conjunction with Excel::Writer::XLSX.
#
# Copyright 2000-2010, John McNamara, jmcnamara@cpan.org
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
use Excel::Writer::XLSX::Utility
  qw(xl_cell_to_rowcol xl_rowcol_to_cell xl_col_to_name xl_range);

our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
our $VERSION = '0.16';


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

    $self->{_sheet_type}  = 0x0200;
    $self->{_orientation} = 0x0;
    $self->{_series}      = [];
    $self->{_embedded}    = 0;

    bless $self, $class;
    $self->_set_default_properties();
    $self->_set_default_config_data();
    return $self;
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

    croak "Must specify 'values' in add_series()" if !exists $arg{values};

    # Parse the ranges to validate them and extract salient information.
    my @value_data    = $self->_parse_series_formula( $arg{values} );
    my @category_data = $self->_parse_series_formula( $arg{categories} );
    my $name_formula  = $self->_parse_series_formula( $arg{name_formula} );

    # Default category count to the same as the value count if not defined.
    if ( !defined $category_data[1] ) {
        $category_data[1] = $value_data[1];
    }

    # Add the parsed data to the user supplied data.
    %arg = (
        @_,
        _values       => \@value_data,
        _categories   => \@category_data,
        _name_formula => $name_formula
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

    $self->{_x_axis_name}    = $arg{name};
    $self->{_x_axis_formula} = $arg{formula};
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

    $self->{_y_axis_name}     = $arg{name};
    $self->{_y_axis_formula}  = $arg{formula};
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

    $self->{_title_name}     = $arg{name};
    $self->{_title_formula}  = $arg{formula};
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

    if ( defined $arg{position} ) {
        if ( lc $arg{position} eq 'none' ) {
            $self->{_legend}->{_visible} = 0;
        }
    }
}


###############################################################################
#
# set_plotarea()
#
# Set the properties of the chart plotarea.
#
sub set_plotarea {

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
# Internal methods. The following section of methods are used for the internal
# structuring of the Chart object and file format.
#
###############################################################################


###############################################################################
#
# _parse_series_formula()
#
# Parse the formula used to define a series. We also extract some range
# information required for _store_series() and the SERIES record.
#
sub _parse_series_formula {

    my $self = shift;

    my $formula  = $_[0];
    my $encoding = 0;
    my $length   = 0;
    my $count    = 0;
    my @tokens;

    return '' if !defined $formula;

    # Strip the = sign at the beginning of the formula string
    $formula =~ s(^=)();

    # Parse the formula using the parser in Formula.pm
    my $parser = $self->{_parser};

    # In order to raise formula errors from the point of view of the calling
    # program we use an eval block and re-raise the error from here.
    #
    eval { @tokens = $parser->parse_formula( $formula ) };

    if ( $@ ) {
        $@ =~ s/\n$//;    # Strip the \n used in the Formula.pm die().
        croak $@;         # Re-raise the error.
    }

    # Force ranges to be a reference class.
    s/_ref3d/_ref3dR/     for @tokens;
    s/_range3d/_range3dR/ for @tokens;
    s/_name/_nameR/       for @tokens;

    # Parse the tokens into a formula string.
    $formula = $parser->parse_tokens( @tokens );

    # Return formula for a single cell as used by title and series name.
    if ( ord $formula == 0x3A ) {
        return $formula;
    }

    # Extract the range from the parse formula.
    if ( ord $formula == 0x3B ) {
        my ( $ptg, $ext_ref, $row_1, $row_2, $col_1, $col_2 ) = unpack 'Cv5',
          $formula;

        # TODO. Remove high bit on relative references.
        $count = $row_2 - $row_1 + 1;
    }

    return ( $formula, $count );
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

    $self->{_legend} = {
        _visible  => 1,
        _position => 0,
        _vertical => 0,
    };

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
# _set_default_config_data()
#
# Setup the default configuration data for a chart.
#
sub _set_default_config_data {

    my $self = shift;

    #<<< Perltidy ignore this.
    $self->{_config} = {
        _axisparent      => [ 0, 0x00F8, 0x01F5, 0x0E7F, 0x0B36              ],
        _axisparent_pos  => [ 2, 2, 0x008C, 0x01AA, 0x0EEA, 0x0C52           ],
        _chart           => [ 0x0000, 0x0000, 0x02DD51E0, 0x01C2B838         ],
        _font_numbers    => [ 5, 10, 0x38B8, 0x22A1, 0x0000                  ],
        _font_series     => [ 6, 10, 0x38B8, 0x22A1, 0x0001                  ],
        _font_title      => [ 7, 12, 0x38B8, 0x22A1, 0x0000                  ],
        _font_axes       => [ 8, 10, 0x38B8, 0x22A1, 0x0001                  ],
        _legend          => [ 0x05F9, 0x0EE9, 0x047D, 0x9C, 0x00, 0x01, 0x0F ],
        _legend_pos      => [ 5, 2, 0x05F9, 0x0EE9, 0, 0                     ],
        _legend_text     => [ 0xFFFFFF46, 0xFFFFFF06, 0, 0, 0x00B1, 0x0000   ],
        _legend_text_pos => [ 2, 2, 0, 0, 0, 0                               ],
        _series_text     => [ 0xFFFFFF46, 0xFFFFFF06, 0, 0, 0x00B1, 0x1020   ],
        _series_text_pos => [ 2, 2, 0, 0, 0, 0                               ],
        _title_text      => [ 0x06E4, 0x0051, 0x01DB, 0x00C4, 0x0081, 0x1030 ],
        _title_text_pos  => [ 2, 2, 0, 0, 0x73, 0x1D                         ],
        _x_axis_text     => [ 0x07E1, 0x0DFC, 0xB2, 0x9C, 0x0081, 0x0000     ],
        _x_axis_text_pos => [ 2, 2, 0, 0,  0x2B,  0x17                       ],
        _y_axis_text     => [ 0x002D, 0x06AA, 0x5F, 0x1CC, 0x0281, 0x00, 90  ],
        _y_axis_text_pos => [ 2, 2, 0, 0, 0x17,  0x44                        ],
    }; #>>>


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


    #<<< Perltidy ignore this.
    $self->{_config} = {
        _axisparent      => [ 0, 0x01D8, 0x031D, 0x0D79, 0x07E9              ],
        _axisparent_pos  => [ 2, 2, 0x010C, 0x0292, 0x0E46, 0x09FD           ],
        _chart           => [ 0x0000, 0x0000, 0x01847FE8, 0x00F47FE8         ],
        _font_numbers    => [ 5, 10, 0x1DC4, 0x1284, 0x0000                  ],
        _font_series     => [ 6, 10, 0x1DC4, 0x1284, 0x0001                  ],
        _font_title      => [ 7, 12, 0x1DC4, 0x1284, 0x0000                  ],
        _font_axes       => [ 8, 10, 0x1DC4, 0x1284, 0x0001                  ],
        _legend          => [ 0x044E, 0x0E4A, 0x088D, 0x0123, 0x0, 0x1, 0xF  ],
        _legend_pos      => [ 5, 2, 0x044E, 0x0E4A, 0, 0                     ],
        _legend_text     => [ 0xFFFFFFD9, 0xFFFFFFC1, 0, 0, 0x00B1, 0x0000   ],
        _legend_text_pos => [ 2, 2, 0, 0, 0, 0                               ],
        _series_text     => [ 0xFFFFFFD9, 0xFFFFFFC1, 0, 0, 0x00B1, 0x1020   ],
        _series_text_pos => [ 2, 2, 0, 0, 0, 0                               ],
        _title_text      => [ 0x060F, 0x004C, 0x038A, 0x016F, 0x0081, 0x1030 ],
        _title_text_pos  => [ 2, 2, 0, 0, 0x73, 0x1D                         ],
        _x_axis_text     => [ 0x07EF, 0x0C8F, 0x153, 0x123, 0x81, 0x00       ],
        _x_axis_text_pos => [ 2, 2, 0, 0, 0x2B, 0x17                         ],
        _y_axis_text     => [ 0x0057, 0x0564, 0xB5, 0x035D, 0x0281, 0x00, 90 ],
        _y_axis_text_pos => [ 2, 2, 0, 0, 0x17, 0x44                         ],
    }; #>>>
}


###############################################################################
#
# XML writing methods.
#
###############################################################################







1;


__END__


=head1 NAME

Chart - A writer class for Excel Charts.

=head1 SYNOPSIS

To create a simple Excel file with a chart using Excel::Writer::XLSX:

    #!/usr/bin/perl -w

    use strict;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'chart.xls' );
    my $worksheet = $workbook->add_worksheet();

    my $chart     = $workbook->add_chart( type => 'column' );

    # Configure the chart.
    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$7',
        values     => '=Sheet1!$B$2:$B$7',
    );

    # Add the worksheet data the chart refers to.
    my $data = [
        [ 'Category', 2, 3, 4, 5, 6, 7 ],
        [ 'Value',    1, 4, 5, 2, 1, 5 ],
    ];

    $worksheet->write( 'A1', $data );

    __END__


=head1 DESCRIPTION

The C<Chart> module is an abstract base class for modules that implement charts in L<Excel::Writer::XLSX>. The information below is applicable to all of the available subclasses.

The C<Chart> module isn't used directly, a chart object is created via the Workbook C<add_chart()> method where the chart type is specified:

    my $chart = $workbook->add_chart( type => 'column' );

Currently the supported chart types are:

=over

=item * C<area>: Creates an Area (filled line) style chart. See L<Excel::Writer::XLSX::Chart::Area>.

=item * C<bar>: Creates a Bar style (transposed histogram) chart. See L<Excel::Writer::XLSX::Chart::Bar>.

=item * C<column>: Creates a column style (histogram) chart. See L<Excel::Writer::XLSX::Chart::Column>.

=item * C<line>: Creates a Line style chart. See L<Excel::Writer::XLSX::Chart::Line>.

=item * C<pie>: Creates an Pie style chart. See L<Excel::Writer::XLSX::Chart::Pie>.

=item * C<scatter>: Creates an Scatter style chart. See L<Excel::Writer::XLSX::Chart::Scatter>.

=item * C<stock>: Creates an Stock style chart. See L<Excel::Writer::XLSX::Chart::Stock>.

=back

More charts and sub-types will be supported in time. See the L</TODO> section.

Methods that are common to all chart types are documented below.

=head1 CHART METHODS

=head2 add_series()

In an Excel chart a "series" is a collection of information such as values, x-axis labels and the name that define which data is plotted. These settings are displayed when you select the C<< Chart -> Source Data... >> menu option.

With a Excel::Writer::XLSX chart object the C<add_series()> method is used to set the properties for a series:

    $chart->add_series(
        categories    => '=Sheet1!$A$2:$A$10',
        values        => '=Sheet1!$B$2:$B$10',
        name          => 'Series name',
        name_formula  => '=Sheet1!$B$1',
    );

The properties that can be set are:

=over

=item * C<values>

This is the most important property of a series and must be set for every chart object. It links the chart with the worksheet data that it displays. Note the format that should be used for the formula. See L</Working with Cell Ranges>.

=item * C<categories>

This sets the chart category labels. The category is more or less the same as the X-axis. In most chart types the C<categories> property is optional and the chart will just assume a sequential series from C<1 .. n>.

=item * C<name>

Set the name for the series. The name is displayed in the chart legend and in the formula bar. The name property is optional and if it isn't supplied will default to C<Series 1 .. n>.

=item * C<name_formula>

Optional, can be used to link the name to a worksheet cell. See L</Chart names and links>.

=back

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

Set the name (title or caption) for the axis. The name is displayed below the X axis. This property is optional. The default is to have no axis name.

=item * C<name_formula>

Optional, can be used to link the name to a worksheet cell. See L</Chart names and links>.

=back

Additional axis properties such as range, divisions and ticks will be made available in later releases. See the L</TODO> section.


=head2 set_y_axis()

The C<set_y_axis()> method is used to set properties of the Y axis.

    $chart->set_y_axis( name => 'Sample weight (kg)' );

The properties that can be set are:

=over

=item * C<name>

Set the name (title or caption) for the axis. The name is displayed to the left of the Y axis. This property is optional. The default is to have no axis name.

=item * C<name_formula>

Optional, can be used to link the name to a worksheet cell. See L</Chart names and links>.

=back

Additional axis properties such as range, divisions and ticks will be made available in later releases. See the L</TODO> section.

=head2 set_title()

The C<set_title()> method is used to set properties of the chart title.

    $chart->set_title( name => 'Year End Results' );

The properties that can be set are:

=over

=item * C<name>

Set the name (title) for the chart. The name is displayed above the chart. This property is optional. The default is to have no chart title.

=item * C<name_formula>

Optional, can be used to link the name to a worksheet cell. See L</Chart names and links>.

=back


=head2 set_legend()

The C<set_legend()> method is used to set properties of the chart legend.

    $chart->set_legend( position => 'none' );

The properties that can be set are:

=over

=item * C<position>

Set the position of the chart legend.

    $chart->set_legend( position => 'none' );

The default legend position is C<bottom>. The currently supported chart positions are:

    none
    bottom

The other legend positions will be added soon.

=back


=head2 set_chartarea()

The C<set_chartarea()> method is used to set the properties of the chart area. In Excel the chart area is the background area behind the chart.

The properties that can be set are:

=over

=item * C<color>

Set the colour of the chart area. The Excel default chart area color is 'white', index 9. See L</Chart object colours>.

=item * C<line_color>

Set the colour of the chart area border line. The Excel default border line colour is 'black', index 9.  See L</Chart object colours>.

=item * C<line_pattern>

Set the pattern of the of the chart area border line. The Excel default pattern is 'none', index 0 for a chart sheet and 'solid', index 1, for an embedded chart. See L</Chart line patterns>.

=item * C<line_weight>

Set the weight of the of the chart area border line. The Excel default weight is 'narrow', index 2. See L</Chart line weights>.

=back

Here is an example of setting several properties:

    $chart->set_chartarea(
        color        => 'red',
        line_color   => 'black',
        line_pattern => 2,
        line_weight  => 3,
    );

Note, for chart sheets the chart area border is off by default. For embedded charts is is on by default.

=head2 set_plotarea()

The C<set_plotarea()> method is used to set properties of the plot area of a chart. In Excel the plot area is the area between the axes on which the chart series are plotted.

The properties that can be set are:

=over

=item * C<visible>

Set the visibility of the plot area. The default is 1 for visible. Set to 0 to hide the plot area and have the same colour as the background chart area.

=item * C<color>

Set the colour of the plot area. The Excel default plot area color is 'silver', index 23. See L</Chart object colours>.

=item * C<line_color>

Set the colour of the plot area border line. The Excel default border line colour is 'gray', index 22. See L</Chart object colours>.

=item * C<line_pattern>

Set the pattern of the of the plot area border line. The Excel default pattern is 'solid', index 1. See L</Chart line patterns>.

=item * C<line_weight>

Set the weight of the of the plot area border line. The Excel default weight is 'narrow', index 2. See L</Chart line weights>.

=back

Here is an example of setting several properties:

    $chart->set_plotarea(
        color        => 'red',
        line_color   => 'black',
        line_pattern => 2,
        line_weight  => 3,
    );



=head1 WORKSHEET METHODS

In Excel a chart sheet (i.e, a chart that isn't embedded) shares properties with data worksheets such as tab selection, headers, footers, margins and print properties.

In Excel::Writer::XLSX you can set chart sheet properties using the same methods that are used for Worksheet objects.

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

    #!/usr/bin/perl -w

    use strict;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'chart_area.xls' );
    my $worksheet = $workbook->add_worksheet();
    my $bold      = $workbook->add_format( bold => 1 );

    # Add the worksheet data that the charts will refer to.
    my $headings = [ 'Number', 'Sample 1', 'Sample 2' ];
    my $data = [
        [ 2, 3, 4, 5, 6, 7 ],
        [ 1, 4, 5, 2, 1, 5 ],
        [ 3, 6, 7, 5, 4, 3 ],
    ];

    $worksheet->write( 'A1', $headings, $bold );
    $worksheet->write( 'A2', $data );

    # Create a new chart object. In this case an embedded chart.
    my $chart = $workbook->add_chart( type => 'area', embedded => 1 );

    # Configure the first series. (Sample 1)
    $chart->add_series(
        name       => 'Sample 1',
        categories => '=Sheet1!$A$2:$A$7',
        values     => '=Sheet1!$B$2:$B$7',
    );

    # Configure the second series. (Sample 2)
    $chart->add_series(
        name       => 'Sample 2',
        categories => '=Sheet1!$A$2:$A$7',
        values     => '=Sheet1!$C$2:$C$7',
    );

    # Add a chart title and some axis labels.
    $chart->set_title ( name => 'Results of sample analysis' );
    $chart->set_x_axis( name => 'Test number' );
    $chart->set_y_axis( name => 'Sample length (cm)' );

    # Insert the chart into the worksheet (with an offset).
    $worksheet->insert_chart( 'D2', $chart, 25, 10 );

    __END__


=begin html

<p>This will produce a chart that looks like this:</p>

<p><center><img src="http://homepage.eircom.net/~jmcnamara/perl/images/area1.jpg" width="527" height="320" alt="Chart example." /></center></p>

=end html


=head1 Chart object colours

Many of the chart objects supported by Spreadsheet::WriteExcl allow the default colours to be changed. Excel provides a palette of 56 colours and in Excel::Writer::XLSX these colours are accessed via their palette index in the range 8..63.

The most commonly used colours can be accessed by name or index.

    black   =>   8,    green    =>  17,    navy     =>  18,
    white   =>   9,    orange   =>  53,    pink     =>  33,
    red     =>  10,    gray     =>  23,    purple   =>  20,
    blue    =>  12,    lime     =>  11,    silver   =>  22,
    yellow  =>  13,    cyan     =>  15,
    brown   =>  16,    magenta  =>  14,

For example the following are equivalent.

    $chart->set_plotarea( color => 10    );
    $chart->set_plotarea( color => 'red' );

The colour palette is shown in C<palette.html> in the C<docs> directory  of the distro. An Excel version of the palette can be generated using C<colors.pl> in the C<examples> directory.

User defined colours can be set using the C<set_custom_color()> workbook method. This and other aspects of using colours are discussed in the "Colours in Excel" section of the main Excel::Writer::XLSX documentation: L<http://search.cpan.org/dist/Spreadsheet-WriteExcel/lib/Spreadsheet/WriteExcel.pm#COLOURS_IN_EXCEL>.

=head1 Chart line patterns

Chart lines patterns can be set using either an index or a name:

    $chart->set_plotarea( weight => 2      );
    $chart->set_plotarea( weight => 'dash' );

Chart lines have 9 possible patterns are follows:

    'none'         => 0,
    'solid'        => 1,
    'dash'         => 2,
    'dot'          => 3,
    'dash-dot'     => 4,
    'dash-dot-dot' => 5,
    'medium-gray'  => 6,
    'dark-gray'    => 7,
    'light-gray'   => 8,

The patterns 1-8 are shown in order in the drop down dialog boxes in Excel. The default pattern is 'solid', index 1.


=head1 Chart line weights

Chart lines weights can be set using either an index or a name:

    $chart->set_plotarea( weight => 1          );
    $chart->set_plotarea( weight => 'hairline' );

Chart lines have 4 possible weights are follows:

    'hairline' => 1,
    'narrow'   => 2,
    'medium'   => 3,
    'wide'     => 4,

The weights 1-4 are shown in order in the drop down dialog boxes in Excel. The default weight is 'narrow', index 2.


=head1 Chart names and links

The C<add_series())>, C<set_x_axis()>, C<set_y_axis()> and C<set_title()> methods all support a C<name> property. In general these names can be either a static string or a link to a worksheet cell. If you choose to use the C<name_formula> property to specify a link then you should also the C<name> property. This isn't strictly required by Excel but some third party applications expect it to be present.

    $chart->set_title(
        name          => 'Year End Results',
        name_formula  => '=Sheet1!$C$1',
    );

These links should be used sparingly since they aren't commonly used in Excel charts.




=head1 Working with Cell Ranges


In the section on C<add_series()> it was noted that the series must be defined using a range formula:

    $chart->add_series( values => '=Sheet1!$B$2:$B$10' );

The worksheet name must be specified (even for embedded charts) and the cell references must be "absolute" references, i.e., they must contain C<$> signs. This is the format that is required by Excel for chart references.

Since it isn't very convenient to work with this type of string programmatically the L<Excel::Writer::XLSX::Utility> module, which is included with Excel::Writer::XLSX, provides a function called C<xl_range_formula()> to convert from zero based row and column cell references to an A1 style formula string.

The syntax is:

    xl_range_formula($sheetname, $row_1, $row_2, $col_1, $col_2)

If you include it in your program, using the standard import syntax, you can use the function as follows:


    # Include the Utility module or just the function you need.
    use Excel::Writer::XLSX::Utility qw( xl_range_formula );
    ...

    # Then use it as required.
    $chart->add_series(
        categories    => xl_range_formula( 'Sheet1', 1, 9, 0, 0 ),
        values        => xl_range_formula( 'Sheet1', 1, 9, 1, 1 ),
    );

    # Which is the same as:
    $chart->add_series(
        categories    => '=Sheet1!$A$2:$A$10',
        values        => '=Sheet1!$B$2:$B$10',
    );

See L<Excel::Writer::XLSX::Utility> for more details.


=head1 TODO

Charts in Excel::Writer::XLSX are a work in progress. More chart types and features will be added in time. Please be patient. Even a small feature can take a week or more to implement, test and document.

Features that are on the TODO list and will be added are:

=over

=item *  Chart sub-types.

=item * Colours and formatting options. For now you will have to make do with the default Excel colours and formats.

=item * Axis controls, gridlines.

=item * 3D charts.

=item * Embedded data in charts for third party application support. See Known Issues.

=item * Additional chart types such as Bubble and Radar. Send an email if you are interested in other types and they will be added to the queue.

=back

If you are interested in sponsoring a feature let me know.

=head1 KNOWN ISSUES

=over

=item * Currently charts don't contain embedded data from which the charts can be rendered. Excel and most other third party applications ignore this and read the data via the links that have been specified. However, some applications may complain or not render charts correctly. The preview option in Mac OS X is an known example. This will be fixed in a later release.

=item * When there are several charts with titles set in a workbook some of the titles may display at a font size of 10 instead of the default 12 until another chart with the title set is viewed.

=item * Stock (and other) charts should have the X-axis dates aligned at an angle for clarity. This will be fixed at a later stage.

=back


=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

