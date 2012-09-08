package Excel::Writer::XLSX::Chart::Scatter;

###############################################################################
#
# Scatter - A class for writing Excel Scatter charts.
#
# Used in conjunction with Excel::Writer::XLSX::Chart.
#
# See formatting note in Excel::Writer::XLSX::Chart.
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
use Excel::Writer::XLSX::Chart;

our @ISA     = qw(Excel::Writer::XLSX::Chart);
our $VERSION = '0.49';


###############################################################################
#
# new()
#
#
sub new {

    my $class = shift;
    my $self  = Excel::Writer::XLSX::Chart->new( @_ );

    $self->{_subtype}          = $self->{_subtype} || 'marker_only';
    $self->{_cross_between}    = 'midCat';
    $self->{_horiz_val_axis}   = 0;
    $self->{_val_axis_postion} = 'b';

    bless $self, $class;
    return $self;
}


##############################################################################
#
# _write_chart_type()
#
# Override the virtual superclass method with a chart specific method.
#
sub _write_chart_type {

    my $self = shift;

    # Write the c:scatterChart element.
    $self->_write_scatter_chart( @_ );
}


##############################################################################
#
# _write_scatter_chart()
#
# Write the <c:scatterChart> element.
#
sub _write_scatter_chart {

    my $self = shift;
    my %args = @_;

    my @series;
    if ( $args{primary_axes} ) {
        @series = $self->_get_primary_axes_series;
    }
    else {
        @series = $self->_get_secondary_axes_series;
    }

    return unless scalar @series;

    my $style   = 'lineMarker';
    my $subtype = $self->{_subtype};

    # Set the user defined chart subtype.
    $style = 'lineMarker'   if $subtype eq 'marker_only';
    $style = 'lineMarker'   if $subtype eq 'straight_with_markers';
    $style = 'lineMarker'   if $subtype eq 'straight';
    $style = 'smoothMarker' if $subtype eq 'smooth_with_markers';
    $style = 'smoothMarker' if $subtype eq 'smooth';

    # Add default formatting to the series data.
    $self->_modify_series_formatting();

    $self->{_writer}->startTag( 'c:scatterChart' );

    # Write the c:scatterStyle element.
    $self->_write_scatter_style( $style );

    # Write the series elements.
    $self->_write_ser( $_ ) for @series;

    # Write the c:marker element.
    $self->_write_marker_value();

    # Write the c:axId elements
    $self->_write_axis_ids( %args );

    $self->{_writer}->endTag( 'c:scatterChart' );
}


##############################################################################
#
# _write_ser()
#
# Over-ridden to write c:xVal/c:yVal instead of c:cat/c:val elements.
#
# Write the <c:ser> element.
#
sub _write_ser {

    my $self   = shift;
    my $series = shift;
    my $index  = $self->{_series_index}++;

    $self->{_writer}->startTag( 'c:ser' );

    # Write the c:idx element.
    $self->_write_idx( $index );

    # Write the c:order element.
    $self->_write_order( $index );

    # Write the series name.
    $self->_write_series_name( $series );

    # Write the c:spPr element.
    $self->_write_sp_pr( $series );

    # Write the c:marker element.
    $self->_write_marker( $series->{_marker} );

    # Write the c:dLbls element.
    $self->_write_d_lbls( $series->{_labels} );

    # Write the c:trendline element.
    $self->_write_trendline( $series->{_trendline} );

    # Write the c:xVal element.
    $self->_write_x_val( $series );

    # Write the c:yVal element.
    $self->_write_y_val( $series );

    # Write the c:smooth element.
    $self->_write_c_smooth();

    $self->{_writer}->endTag( 'c:ser' );
}


##############################################################################
#
# _write_plot_area()
#
# Over-ridden to have 2 valAx elements for scatter charts instead of
# catAx/valAx.
#
# Write the <c:plotArea> element.
#
sub _write_plot_area {

    my $self = shift;

    $self->{_writer}->startTag( 'c:plotArea' );

    # Write the c:layout element.
    $self->_write_layout();

    # Write the subclass chart type elements for primary and secondary axes
    $self->_write_chart_type( primary_axes => 1 );
    $self->_write_chart_type( primary_axes => 0 );

    # Write c:catAx and c:valAx elements for series using primary axes
    $self->_write_cat_val_axis(
        x_axis   => $self->{_x_axis},
        y_axis   => $self->{_y_axis},
        axis_ids => $self->{_axis_ids},
        position => 'b',
    );
    my $tmp = $self->{_horiz_val_axis};
    $self->{_horiz_val_axis} = 1;
    $self->_write_val_axis(
        x_axis   => $self->{_x_axis},
        y_axis   => $self->{_y_axis},
        axis_ids => $self->{_axis_ids},
        position => 'l',
    );
    $self->{_horiz_val_axis} = $tmp;

    # Write c:valAx and c:catAx elements for series using secondary axes
    $self->_write_cat_val_axis(
        x_axis   => $self->{_x2_axis},
        y_axis   => $self->{_y2_axis},
        axis_ids => $self->{_axis2_ids},
        position => 'b',
    );
    $self->{_horiz_val_axis} = 1;
    $self->_write_val_axis(
        x_axis   => $self->{_x2_axis},
        y_axis   => $self->{_y2_axis},
        axis_ids => $self->{_axis2_ids},
        position => 'l',
    );

    $self->{_writer}->endTag( 'c:plotArea' );
}


##############################################################################
#
# _write_x_val()
#
# Write the <c:xVal> element.
#
sub _write_x_val {

    my $self    = shift;
    my $series  = shift;
    my $formula = $series->{_categories};
    my $data_id = $series->{_cat_data_id};
    my $data    = $self->{_formula_data}->[$data_id];

    $self->{_writer}->startTag( 'c:xVal' );

    # Check the type of cached data.
    my $type = $self->_get_data_type( $data );

    # TODO. Can a scatter plot have non-numeric data.

    if ( $type eq 'str' ) {

        # Write the c:numRef element.
        $self->_write_str_ref( $formula, $data, $type );
    }
    else {

        # Write the c:numRef element.
        $self->_write_num_ref( $formula, $data, $type );
    }

    $self->{_writer}->endTag( 'c:xVal' );
}


##############################################################################
#
# _write_y_val()
#
# Write the <c:yVal> element.
#
sub _write_y_val {

    my $self    = shift;
    my $series  = shift;
    my $formula = $series->{_values};
    my $data_id = $series->{_val_data_id};
    my $data    = $self->{_formula_data}->[$data_id];

    $self->{_writer}->startTag( 'c:yVal' );

    # Unlike Cat axes data should only be numeric.

    # Write the c:numRef element.
    $self->_write_num_ref( $formula, $data, 'num' );

    $self->{_writer}->endTag( 'c:yVal' );
}


##############################################################################
#
# _write_scatter_style()
#
# Write the <c:scatterStyle> element.
#
sub _write_scatter_style {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:scatterStyle', @attributes );
}


##############################################################################
#
# _write_c_smooth()
#
# Write the <c:smooth> element.
#
sub _write_c_smooth {

    my $self    = shift;
    my $subtype = $self->{_subtype};
    my $val     = 1;

    return unless $subtype =~ /smooth/;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:smooth', @attributes );
}


##############################################################################
#
# _modify_series_formatting()
#
# Add default formatting to the series data unless it has already been
# specified by the user.
#
sub _modify_series_formatting {

    my $self    = shift;
    my $subtype = $self->{_subtype};

    # The default scatter style "markers only" requires a line type.
    if ( $subtype eq 'marker_only' ) {

        # Go through each series and define default values.
        for my $series ( @{ $self->{_series} } ) {

            # Set a line type unless there is already a user defined type.
            if ( !$series->{_line}->{_defined} ) {
                $series->{_line} = {
                    width    => 2.25,
                    none     => 1,
                    _defined => 1,
                };
            }
        }
    }

    # Turn markers off for subtypes that don't have them.
    if ( $subtype !~ /marker/ ) {

        # Go through each series and define default values.
        for my $series ( @{ $self->{_series} } ) {

            # Set a marker type unless there is already a user defined type.
            if ( !$series->{_marker}->{_defined} ) {
                $series->{_marker} = {
                    type     => 'none',
                    _defined => 1,
                };
            }
        }
    }

}


1;


__END__


=head1 NAME

Scatter - A class for writing Excel Scatter charts.

=head1 SYNOPSIS

To create a simple Excel file with a Scatter chart using Excel::Writer::XLSX:

    #!/usr/bin/perl

    use strict;
    use warnings;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'chart.xlsx' );
    my $worksheet = $workbook->add_worksheet();

    my $chart     = $workbook->add_chart( type => 'scatter' );

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

This module implements Scatter charts for L<Excel::Writer::XLSX>. The chart object is created via the Workbook C<add_chart()> method:

    my $chart = $workbook->add_chart( type => 'scatter' );

Once the object is created it can be configured via the following methods that are common to all chart classes:

    $chart->add_series();
    $chart->set_x_axis();
    $chart->set_y_axis();
    $chart->set_title();

These methods are explained in detail in L<Excel::Writer::XLSX::Chart>. Class specific methods or settings, if any, are explained below.

=head1 Scatter Chart Methods

The C<Scatter> chart module also supports the following sub-types:

    markers_only (the default)
    straight_with_markers
    straight
    smooth_with_markers
    smooth

These can be specified at creation time via the C<add_chart()> Worksheet method:

    my $chart = $workbook->add_chart(
        type    => 'scatter',
        subtype => 'straight_with_markers'
    );

=head1 EXAMPLE

Here is a complete example that demonstrates most of the available features when creating a chart.

    #!/usr/bin/perl

    use strict;
    use warnings;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'chart_scatter.xlsx' );
    my $worksheet = $workbook->add_worksheet();
    my $bold      = $workbook->add_format( bold => 1 );

    # Add the worksheet data that the charts will refer to.
    my $headings = [ 'Number', 'Batch 1', 'Batch 2' ];
    my $data = [
        [ 2, 3, 4, 5, 6, 7 ],
        [ 10, 40, 50, 20, 10, 50 ],
        [ 30, 60, 70, 50, 40, 30 ],

    ];

    $worksheet->write( 'A1', $headings, $bold );
    $worksheet->write( 'A2', $data );

    # Create a new chart object. In this case an embedded chart.
    my $chart = $workbook->add_chart( type => 'scatter', embedded => 1 );

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

    # Set an Excel chart style. Colors with white outline and shadow.
    $chart->set_style( 10 );

    # Insert the chart into the worksheet (with an offset).
    $worksheet->insert_chart( 'D2', $chart, 25, 10 );

    __END__


=begin html

<p>This will produce a chart that looks like this:</p>

<p><center><img src="http://homepage.eircom.net/~jmcnamara/perl/images/2007/scatter1.jpg" width="483" height="291" alt="Chart example." /></center></p>

=end html


=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MM-MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

