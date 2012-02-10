package Excel::Writer::XLSX::Chart::Pie;

###############################################################################
#
# Pie - A class for writing Excel Pie charts.
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
our $VERSION = '0.46';


###############################################################################
#
# new()
#
#
sub new {

    my $class = shift;
    my $self  = Excel::Writer::XLSX::Chart->new( @_ );

    $self->{_vary_data_color} = 1;

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

    # Write the c:pieChart element.
    $self->_write_pie_chart();
}


##############################################################################
#
# _write_pie_chart()
#
# Write the <c:pieChart> element.
#
sub _write_pie_chart {

    my $self = shift;

    $self->{_writer}->startTag( 'c:pieChart' );

    # Write the c:varyColors element.
    $self->_write_vary_colors();

    # Write the series elements.
    $self->_write_series();

   # Write the c:firstSliceAng element.
    $self->_write_first_slice_ang();

    $self->{_writer}->endTag( 'c:pieChart' );
}


##############################################################################
#
# _write_plot_area().
#
# Over-ridden method to remove the cat_axis() and val_axis() code since
# Pie charts don't require those axes.
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

    $self->{_writer}->endTag( 'c:plotArea' );
}


##############################################################################
#
# _write_series().
#
# Over-ridden method to remove axis_id code since Pie charts  don't require
# val and cat axes.
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
}



##############################################################################
#
# _write_legend().
#
# Over-ridden method to add <c:txPr> to legend.
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

    # Write the c:txPr element. Over-ridden.
    $self->_write_tx_pr_legend();

    $self->{_writer}->endTag( 'c:legend' );
}



##############################################################################
#
# _write_tx_pr_legend()
#
# Write the <c:txPr> element for legends.
#
sub _write_tx_pr_legend {

    my $self  = shift;
    my $horiz = 0;

    $self->{_writer}->startTag( 'c:txPr' );

    # Write the a:bodyPr element.
    $self->_write_a_body_pr( $horiz );

    # Write the a:lstStyle element.
    $self->_write_a_lst_style();

    # Write the a:p element.
    $self->_write_a_p_legend();

    $self->{_writer}->endTag( 'c:txPr' );
}


##############################################################################
#
# _write_a_p_legend()
#
# Write the <a:p> element for legends.
#
sub _write_a_p_legend {

    my $self  = shift;
    my $title = shift;

    $self->{_writer}->startTag( 'a:p' );

    # Write the a:pPr element.
    $self->_write_a_p_pr_legend();

    # Write the a:endParaRPr element.
    $self->_write_a_end_para_rpr();

    $self->{_writer}->endTag( 'a:p' );
}


##############################################################################
#
# _write_a_p_pr_legend()
#
# Write the <a:pPr> element for legends.
#
sub _write_a_p_pr_legend {

    my $self = shift;
    my $rtl  = 0;

    my @attributes = ( 'rtl' => $rtl );

    $self->{_writer}->startTag( 'a:pPr', @attributes );

    # Write the a:defRPr element.
    $self->_write_a_def_rpr();

    $self->{_writer}->endTag( 'a:pPr' );
}


##############################################################################
#
# _write_vary_colors()
#
# Write the <c:varyColors> element.
#
sub _write_vary_colors {

    my $self = shift;
    my $val  = 1;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:varyColors', @attributes );
}


##############################################################################
#
# _write_first_slice_ang()
#
# Write the <c:firstSliceAng> element.
#
sub _write_first_slice_ang {

    my $self = shift;
    my $val  = 0;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'c:firstSliceAng', @attributes );
}

1;


__END__


=head1 NAME

Pie - A class for writing Excel Pie charts.

=head1 SYNOPSIS

To create a simple Excel file with a Pie chart using Excel::Writer::XLSX:

    #!/usr/bin/perl

    use strict;
    use warnings;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'chart.xlsx' );
    my $worksheet = $workbook->add_worksheet();

    my $chart     = $workbook->add_chart( type => 'pie' );

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

This module implements Pie charts for L<Excel::Writer::XLSX>. The chart object is created via the Workbook C<add_chart()> method:

    my $chart = $workbook->add_chart( type => 'pie' );

Once the object is created it can be configured via the following methods that are common to all chart classes:

    $chart->add_series();
    $chart->set_title();

These methods are explained in detail in L<Excel::Writer::XLSX::Chart>. Class specific methods or settings, if any, are explained below.

=head1 Pie Chart Methods

There aren't currently any pie chart specific methods. See the TODO section of L<Excel::Writer::XLSX::Chart>.

A Pie chart doesn't have an X or Y axis so the following common chart methods are ignored.

    $chart->set_x_axis();
    $chart->set_y_axis();

=head1 EXAMPLE

Here is a complete example that demonstrates most of the available features when creating a chart.

    #!/usr/bin/perl

    use strict;
    use warnings;
    use Excel::Writer::XLSX;

    my $workbook  = Excel::Writer::XLSX->new( 'chart_pie.xlsx' );
    my $worksheet = $workbook->add_worksheet();
    my $bold      = $workbook->add_format( bold => 1 );

    # Add the worksheet data that the charts will refer to.
    my $headings = [ 'Category', 'Values' ];
    my $data = [
        [ 'Apple', 'Cherry', 'Pecan' ],
        [ 60,       30,       10     ],
    ];

    $worksheet->write( 'A1', $headings, $bold );
    $worksheet->write( 'A2', $data );

    # Create a new chart object. In this case an embedded chart.
    my $chart = $workbook->add_chart( type => 'pie', embedded => 1 );

    # Configure the series. Note the use of the array ref to define ranges:
    # [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
    $chart->add_series(
        name       => 'Pie sales data',
        categories => [ 'Sheet1', 1, 3, 0, 0 ],
        values     => [ 'Sheet1', 1, 3, 1, 1 ],
    );

    # Add a title.
    $chart->set_title( name => 'Popular Pie Types' );

    # Set an Excel chart style. Colors with white outline and shadow.
    $chart->set_style( 10 );

    # Insert the chart into the worksheet (with an offset).
    $worksheet->insert_chart( 'C2', $chart, 25, 10 );

    __END__


=begin html

<p>This will produce a chart that looks like this:</p>

<p><center><img src="http://homepage.eircom.net/~jmcnamara/perl/images/2007/pie1.jpg" width="483" height="291" alt="Chart example." /></center></p>

=end html


=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MM-MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

