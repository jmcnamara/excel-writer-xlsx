package Excel::Writer::XLSX::Chart::Stock;

###############################################################################
#
# Stock - A writer class for Excel Stock charts.
#
# Used in conjunction with Excel::Writer::XLSX::Chart.
#
# See formatting note in Excel::Writer::XLSX::Chart.
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
use Excel::Writer::XLSX::Chart;

our @ISA     = qw(Excel::Writer::XLSX::Chart);
our $VERSION = '0.16';


###############################################################################
#
# new()
#
#
sub new {

    my $class = shift;
    my $self  = Excel::Writer::XLSX::Chart->new( @_ );

    $self->{_default_marker} = 'none';

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

    # Write the c:stockChart element.
    $self->_write_stock_chart();
}


##############################################################################
#
# _write_stock_chart()
#
# Write the <c:stockChart> element.
#
sub _write_stock_chart {

    my $self = shift;

    $self->{_writer}->startTag( 'c:stockChart' );

    # Write the series elements.
    $self->_write_series();

    $self->{_writer}->endTag( 'c:stockChart' );
}


##############################################################################
#
# _write_series()
#
# Over-ridden to add hi_low_lines(). TODO. Refactor up into the SUPER class.
#
# Write the series elements.
#
sub _write_series {

    my $self = shift;

    # Write each series with subelements.
    my $index = 0;
    for my $series ( @{ $self->{_series} } ) {
        if ( $index == 2 ) {
            $self->{_default_marker} = 'dot';
        }

        $self->_write_ser( $index++, $series );
    }

    # Write the c:hiLowLines element.
    $self->_write_hi_low_lines();

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
    my $categories = $series->{_categories};
    my $values     = $series->{_values};

    $self->{_writer}->startTag( 'c:ser' );

    # Write the c:idx element.
    $self->_write_idx( $index );

    # Write the c:order element.
    $self->_write_order( $index );

    # Write the series name.
    $self->_write_series_name( $series );

    # Write the c:spPr element.
    $self->_write_sp_pr();

    # Write the c:marker element.
    $self->_write_marker( );

    # Write the c:cat element.
    $self->_write_cat( $categories );

    # Write the c:val element.
    $self->_write_val( $values );

    $self->{_writer}->endTag( 'c:ser' );
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

    # Write the c:dateAx element.
    $self->_write_date_axis();

    # Write the c:catAx element.
    $self->_write_val_axis();

    $self->{_writer}->endTag( 'c:plotArea' );
}


1;


__END__


=head1 NAME

Stock - A writer class for Excel Stock charts.

=head1 SYNOPSIS

To create a simple Excel file with a Stock chart using Excel::Writer::XLSX:

    #!/usr/bin/perl -w

    use strict;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'chart.xls' );
    my $worksheet = $workbook->add_worksheet();

    my $chart     = $workbook->add_chart( type => 'stock' );

    # Add a series for each Open-High-Low-Close.
    $chart->add_series( categories => '=Sheet1!$A$2:$A$6', values => '=Sheet1!$B$2:$B$6' );
    $chart->add_series( categories => '=Sheet1!$A$2:$A$6', values => '=Sheet1!$C$2:$C$6' );
    $chart->add_series( categories => '=Sheet1!$A$2:$A$6', values => '=Sheet1!$D$2:$D$6' );
    $chart->add_series( categories => '=Sheet1!$A$2:$A$6', values => '=Sheet1!$E$2:$E$6' );

    # Add the worksheet data the chart refers to.
    # ... See the full example below.

    __END__


=head1 DESCRIPTION

This module implements Stock charts for L<Excel::Writer::XLSX>. The chart object is created via the Workbook C<add_chart()> method:

    my $chart = $workbook->add_chart( type => 'stock' );

Once the object is created it can be configured via the following methods that are common to all chart classes:

    $chart->add_series();
    $chart->set_x_axis();
    $chart->set_y_axis();
    $chart->set_title();

These methods are explained in detail in L<Excel::Writer::XLSX::Chart>. Class specific methods or settings, if any, are explained below.

=head1 Stock Chart Methods

There aren't currently any stock chart specific methods. See the TODO section of L<Excel::Writer::XLSX::Chart>.

The default Stock chart is an Open-High-Low-Close chart. A series must be added for each of these data sources.

The default Stock chart is in black and white. User defined colours will be added at a later stage.

=head1 EXAMPLE

Here is a complete example that demonstrates most of the available features when creating a Stock chart.

    #!/usr/bin/perl -w

    use strict;
    use Excel::Writer::XLSX;

    my $workbook    = Excel::Writer::XLSX->new( 'chart_stock_ex.xls' );
    my $worksheet   = $workbook->add_worksheet();
    my $bold        = $workbook->add_format( bold => 1 );
    my $date_format = $workbook->add_format( num_format => 'dd/mm/yyyy' );

    # Add the worksheet data that the charts will refer to.
    my $headings = [ 'Date', 'Open', 'High', 'Low', 'Close' ];
    my @data = (
        [ '2009-08-23', 110.75, 113.48, 109.05, 109.40 ],
        [ '2009-08-24', 111.24, 111.60, 103.57, 104.87 ],
        [ '2009-08-25', 104.96, 108.00, 103.88, 106.00 ],
        [ '2009-08-26', 104.95, 107.95, 104.66, 107.91 ],
        [ '2009-08-27', 108.10, 108.62, 105.69, 106.15 ],
    );

    $worksheet->write( 'A1', $headings, $bold );

    my $row = 1;
    for my $data ( @data ) {
        $worksheet->write( $row, 0, $data->[0], $date_format );
        $worksheet->write( $row, 1, $data->[1] );
        $worksheet->write( $row, 2, $data->[2] );
        $worksheet->write( $row, 3, $data->[3] );
        $worksheet->write( $row, 4, $data->[4] );
        $row++;
    }

    # Create a new chart object. In this case an embedded chart.
    my $chart = $workbook->add_chart( type => 'stock', embedded => 1 );

    # Add a series for each of the Open-High-Low-Close columns.
    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$6',
        values     => '=Sheet1!$B$2:$B$6',
        name       => 'Open',
    );

    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$6',
        values     => '=Sheet1!$C$2:$C$6',
        name       => 'High',
    );

    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$6',
        values     => '=Sheet1!$D$2:$D$6',
        name       => 'Low',
    );

    $chart->add_series(
        categories => '=Sheet1!$A$2:$A$6',
        values     => '=Sheet1!$E$2:$E$6',
        name       => 'Close',
    );

    # Add a chart title and some axis labels.
    $chart->set_title( name => 'Open-High-Low-Close', );
    $chart->set_x_axis( name => 'Date', );
    $chart->set_y_axis( name => 'Share price', );

    # Insert the chart into the worksheet (with an offset).
    $worksheet->insert_chart( 'F2', $chart, 25, 10 );

    __END__


=begin html

<p>This will produce a chart that looks like this:</p>

<p><center><img src="http://homepage.eircom.net/~jmcnamara/perl/images/stock1.jpg" width="527" height="320" alt="Chart example." /></center></p>

=end html


=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

