package Excel::Writer::XLSX::Worksheet;

###############################################################################
#
# Worksheet - A class for writing Excel Worksheets.
#
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
use File::Temp 'tempfile';
use Excel::Writer::XLSX::Format;
use Excel::Writer::XLSX::Drawing;
use Excel::Writer::XLSX::Package::XMLwriter;
use Excel::Writer::XLSX::Utility
  qw(xl_cell_to_rowcol xl_rowcol_to_cell xl_col_to_name xl_range);

our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
our $VERSION = '0.46';


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

    my $class  = shift;
    my $self   = Excel::Writer::XLSX::Package::XMLwriter->new();
    my $rowmax = 1_048_576;
    my $colmax = 16_384;
    my $strmax = 32767;

    $self->{_name}         = $_[0];
    $self->{_index}        = $_[1];
    $self->{_activesheet}  = $_[2];
    $self->{_firstsheet}   = $_[3];
    $self->{_str_total}    = $_[4];
    $self->{_str_unique}   = $_[5];
    $self->{_str_table}    = $_[6];
    $self->{_1904}         = $_[7];
    $self->{_palette}      = $_[8];
    $self->{_optimization} = $_[9] || 0;
    $self->{_tempdir}      = $_[10];

    $self->{_ext_sheets} = [];
    $self->{_fileclosed} = 0;

    $self->{_xls_rowmax} = $rowmax;
    $self->{_xls_colmax} = $colmax;
    $self->{_xls_strmax} = $strmax;
    $self->{_dim_rowmin} = undef;
    $self->{_dim_rowmax} = undef;
    $self->{_dim_colmin} = undef;
    $self->{_dim_colmax} = undef;

    $self->{_colinfo}    = [];
    $self->{_selections} = [];
    $self->{_hidden}     = 0;
    $self->{_active}     = 0;
    $self->{_tab_color}  = 0;

    $self->{_panes}       = [];
    $self->{_active_pane} = 3;
    $self->{_selected}    = 0;

    $self->{_page_setup_changed} = 0;
    $self->{_paper_size}         = 0;
    $self->{_orientation}        = 1;

    $self->{_print_options_changed} = 0;
    $self->{_hcenter}               = 0;
    $self->{_vcenter}               = 0;
    $self->{_print_gridlines}       = 0;
    $self->{_screen_gridlines}      = 1;
    $self->{_print_headers}         = 0;

    $self->{_header_footer_changed} = 0;
    $self->{_header}                = '';
    $self->{_footer}                = '';

    $self->{_margin_left}   = 0.7;
    $self->{_margin_right}  = 0.7;
    $self->{_margin_top}    = 0.75;
    $self->{_margin_bottom} = 0.75;
    $self->{_margin_header} = 0.3;
    $self->{_margin_footer} = 0.3;

    $self->{_repeat_rows} = '';
    $self->{_repeat_cols} = '';
    $self->{_print_area}  = '';

    $self->{_page_order}     = 0;
    $self->{_black_white}    = 0;
    $self->{_draft_quality}  = 0;
    $self->{_print_comments} = 0;
    $self->{_page_start}     = 0;

    $self->{_fit_page}   = 0;
    $self->{_fit_width}  = 0;
    $self->{_fit_height} = 0;

    $self->{_hbreaks} = [];
    $self->{_vbreaks} = [];

    $self->{_protect}  = 0;
    $self->{_password} = undef;

    $self->{_set_cols} = {};
    $self->{_set_rows} = {};

    $self->{_zoom}              = 100;
    $self->{_zoom_scale_normal} = 1;
    $self->{_print_scale}       = 100;
    $self->{_right_to_left}     = 0;
    $self->{_show_zeros}        = 1;
    $self->{_leading_zeros}     = 0;

    $self->{_outline_row_level} = 0;
    $self->{_outline_col_level} = 0;
    $self->{_outline_style}     = 0;
    $self->{_outline_below}     = 1;
    $self->{_outline_right}     = 1;
    $self->{_outline_on}        = 1;
    $self->{_outline_changed}   = 0;

    $self->{_names} = {};

    $self->{_write_match} = [];

    $self->{prev_col} = -1;

    $self->{_table} = [];
    $self->{_merge} = [];

    $self->{_has_comments}     = 0;
    $self->{_comments}         = {};
    $self->{_comments_array}   = [];
    $self->{_comments_author}  = '';
    $self->{_comments_visible} = 0;
    $self->{_vml_shape_id}     = 1024;

    $self->{_autofilter}   = '';
    $self->{_filter_on}    = 0;
    $self->{_filter_range} = [];
    $self->{_filter_cols}  = {};

    $self->{_col_sizes}        = {};
    $self->{_row_sizes}        = {};
    $self->{_col_formats}      = {};
    $self->{_col_size_changed} = 0;
    $self->{_row_size_changed} = 0;

    $self->{_hlink_count}            = 0;
    $self->{_hlink_refs}             = [];
    $self->{_external_hyper_links}   = [];
    $self->{_external_drawing_links} = [];
    $self->{_external_comment_links} = [];
    $self->{_drawing_links}          = [];
    $self->{_charts}                 = [];
    $self->{_images}                 = [];
    $self->{_drawing}                = 0;

    $self->{_rstring}      = '';
    $self->{_previous_row} = 0;

    if ( $self->{_optimization} == 1 ) {
        my $fh  = tempfile( DIR => $self->{_tempdir} );
        binmode $fh, ':utf8';

        my $writer = Excel::Writer::XLSX::Package::XMLwriterSimple->new( $fh );

        $self->{_cell_data_fh} = $fh;
        $self->{_writer} = $writer;
    }

    $self->{_validations}  = [];
    $self->{_cond_formats} = {};
    $self->{_dxf_priority} = 1;

    bless $self, $class;
    return $self;
}

###############################################################################
#
# _set_xml_writer()
#
# Over-ridden to ensure that write_single_row() is called for the final row
# when optimisation mode is on.
#
sub _set_xml_writer {

    my $self     = shift;
    my $filename = shift;

    if ( $self->{_optimization} == 1 ) {
        $self->_write_single_row();
    }

    $self->SUPER::_set_xml_writer( $filename );
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

    # Write the root worksheet element.
    $self->_write_worksheet();

    # Write the worksheet properties.
    $self->_write_sheet_pr();

    # Write the worksheet dimensions.
    $self->_write_dimension();

    # Write the sheet view properties.
    $self->_write_sheet_views();

    # Write the sheet format properties.
    $self->_write_sheet_format_pr();

    # Write the sheet column info.
    $self->_write_cols();

    # Write the worksheet data such as rows columns and cells.
    if ( $self->{_optimization} == 0 ) {
        $self->_write_sheet_data();
    }
    else {
        $self->_write_optimized_sheet_data();
    }


    # Write the sheetProtection element.
    $self->_write_sheet_protection();

    # Write the worksheet calculation properties.
    #$self->_write_sheet_calc_pr();

    # Write the worksheet phonetic properties.
    #$self->_write_phonetic_pr();

    # Write the autoFilter element.
    $self->_write_auto_filter();

    # Write the mergeCells element.
    $self->_write_merge_cells();

    # Write the conditional formats.
    $self->_write_conditional_formats();

    # Write the dataValidations element.
    $self->_write_data_validations();

    # Write the hyperlink element.
    $self->_write_hyperlinks();

    # Write the printOptions element.
    $self->_write_print_options();

    # Write the worksheet page_margins.
    $self->_write_page_margins();

    # Write the worksheet page setup.
    $self->_write_page_setup();

    # Write the headerFooter element.
    $self->_write_header_footer();

    # Write the rowBreaks element.
    $self->_write_row_breaks();

    # Write the colBreaks element.
    $self->_write_col_breaks();

    # Write the drawing element.
    $self->_write_drawings();

    # Write the legacyDrawing element.
    $self->_write_legacy_drawing();

    # Write the worksheet extension storage.
    #$self->_write_ext_lst();

    # Close the worksheet tag.
    $self->{_writer}->endTag( 'worksheet' );

    # Close the XML writer object and filehandle.
    $self->{_writer}->end();
    $self->{_writer}->getOutput()->close();
}


###############################################################################
#
# _close()
#
# Write the worksheet elements.
#
sub _close {

    # TODO. Unused. Remove after refactoring.
    my $self       = shift;
    my $sheetnames = shift;
    my $num_sheets = scalar @$sheetnames;
}


###############################################################################
#
# get_name().
#
# Retrieve the worksheet name.
#
sub get_name {

    my $self = shift;

    return $self->{_name};
}


###############################################################################
#
# select()
#
# Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
# highlighted.
#
sub select {

    my $self = shift;

    $self->{_hidden}   = 0;    # Selected worksheet can't be hidden.
    $self->{_selected} = 1;
}


###############################################################################
#
# activate()
#
# Set this worksheet as the active worksheet, i.e. the worksheet that is
# displayed when the workbook is opened. Also set it as selected.
#
sub activate {

    my $self = shift;

    $self->{_hidden}   = 0;    # Active worksheet can't be hidden.
    $self->{_selected} = 1;
    ${ $self->{_activesheet} } = $self->{_index};
}


###############################################################################
#
# hide()
#
# Hide this worksheet.
#
sub hide {

    my $self = shift;

    $self->{_hidden} = 1;

    # A hidden worksheet shouldn't be active or selected.
    $self->{_selected} = 0;
    ${ $self->{_activesheet} } = 0;
    ${ $self->{_firstsheet} }  = 0;
}


###############################################################################
#
# set_first_sheet()
#
# Set this worksheet as the first visible sheet. This is necessary
# when there are a large number of worksheets and the activated
# worksheet is not visible on the screen.
#
sub set_first_sheet {

    my $self = shift;

    $self->{_hidden} = 0;    # Active worksheet can't be hidden.
    ${ $self->{_firstsheet} } = $self->{_index};
}


###############################################################################
#
# protect( $password )
#
# Set the worksheet protection flags to prevent modification of worksheet
# objects.
#
sub protect {

    my $self     = shift;
    my $password = shift || '';
    my $options  = shift || {};

    if ( $password ne '' ) {
        $password = $self->_encode_password( $password );
    }

    # Default values for objects that can be protected.
    my %defaults = (
        sheet                 => 1,
        content               => 0,
        objects               => 0,
        scenarios             => 0,
        format_cells          => 0,
        format_columns        => 0,
        format_rows           => 0,
        insert_columns        => 0,
        insert_rows           => 0,
        insert_hyperlinks     => 0,
        delete_columns        => 0,
        delete_rows           => 0,
        select_locked_cells   => 1,
        sort                  => 0,
        autofilter            => 0,
        pivot_tables          => 0,
        select_unlocked_cells => 1,
    );


    # Overwrite the defaults with user specified values.
    for my $key ( keys %{$options} ) {

        if ( exists $defaults{$key} ) {
            $defaults{$key} = $options->{$key};
        }
        else {
            carp "Unknown protection object: $key\n";
        }
    }

    # Set the password after the user defined values.
    $defaults{password} = $password;

    $self->{_protect} = \%defaults;
}


###############################################################################
#
# _encode_password($password)
#
# Based on the algorithm provided by Daniel Rentz of OpenOffice.
#
sub _encode_password {

    use integer;

    my $self      = shift;
    my $plaintext = $_[0];
    my $password;
    my $count;
    my @chars;
    my $i = 0;

    $count = @chars = split //, $plaintext;

    foreach my $char ( @chars ) {
        my $low_15;
        my $high_15;
        $char    = ord( $char ) << ++$i;
        $low_15  = $char & 0x7fff;
        $high_15 = $char & 0x7fff << 15;
        $high_15 = $high_15 >> 15;
        $char    = $low_15 | $high_15;
    }

    $password = 0x0000;
    $password ^= $_ for @chars;
    $password ^= $count;
    $password ^= 0xCE4B;

    return sprintf "%X", $password;
}


###############################################################################
#
# set_column($firstcol, $lastcol, $width, $format, $hidden, $level)
#
# Set the width of a single column or a range of columns.
# See also: _write_col_info
#
sub set_column {

    my $self = shift;
    my @data = @_;
    my $cell = $data[0];

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $cell =~ /^\D/ ) {
        @data = $self->_substitute_cellref( @_ );

        # Returned values $row1 and $row2 aren't required here. Remove them.
        shift @data;    # $row1
        splice @data, 1, 1;    # $row2
    }

    return if @data < 3;       # Ensure at least $firstcol, $lastcol and $width
    return if not defined $data[0];    # Columns must be defined.
    return if not defined $data[1];

    # Assume second column is the same as first if 0. Avoids KB918419 bug.
    $data[1] = $data[0] if $data[1] == 0;

    # Ensure 2nd col is larger than first. Also for KB918419 bug.
    ( $data[0], $data[1] ) = ( $data[1], $data[0] ) if $data[0] > $data[1];


    # Check that cols are valid and store max and min values with default row.
    # NOTE: The check shouldn't modify the row dimensions and should only modify
    #       the column dimensions in certain cases.
    my $ignore_row = 1;
    my $ignore_col = 1;
    $ignore_col = 0 if ref $data[3];          # Column has a format.
    $ignore_col = 0 if $data[2] && $data[4];  # Column has a width but is hidden

    return -2
      if $self->_check_dimensions( 0, $data[0], $ignore_row, $ignore_col );
    return -2
      if $self->_check_dimensions( 0, $data[1], $ignore_row, $ignore_col );

    # Set the limits for the outline levels (0 <= x <= 7).
    $data[5] = 0 unless defined $data[5];
    $data[5] = 0 if $data[5] < 0;
    $data[5] = 7 if $data[5] > 7;

    if ( $data[5] > $self->{_outline_col_level} ) {
        $self->{_outline_col_level} = $data[5];
    }

    # Store the column data.
    push @{ $self->{_colinfo} }, [@data];

    # Store the column change to allow optimisations.
    $self->{_col_size_changed} = 1;

    # Store the col sizes for use when calculating image vertices taking
    # hidden columns into account. Also store the column formats.
    my $width = $data[4] ? 0 : $data[2];    # Set width to zero if col is hidden
    $width ||= 0;                           # Ensure width isn't undef.
    my $format = $data[3];

    my ( $firstcol, $lastcol ) = @data;

    foreach my $col ( $firstcol .. $lastcol ) {
        $self->{_col_sizes}->{$col} = $width;
        $self->{_col_formats}->{$col} = $format if $format;
    }
}


###############################################################################
#
# set_selection()
#
# Set which cell or cells are selected in a worksheet.
#
sub set_selection {

    my $self = shift;
    my $pane;
    my $active_cell;
    my $sqref;

    return unless @_;

    # Check for a cell reference in A1 notation and substitute row and column.
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }


    # There should be either 2 or 4 arguments.
    if ( @_ == 2 ) {

        # Single cell selection.
        $active_cell = xl_rowcol_to_cell( $_[0], $_[1] );
        $sqref = $active_cell;
    }
    elsif ( @_ == 4 ) {

        # Range selection.
        $active_cell = xl_rowcol_to_cell( $_[0], $_[1] );

        my ( $row_first, $col_first, $row_last, $col_last ) = @_;

        # Swap last row/col for first row/col as necessary
        if ( $row_first > $row_last ) {
            ( $row_first, $row_last ) = ( $row_last, $row_first );
        }

        if ( $col_first > $col_last ) {
            ( $col_first, $col_last ) = ( $col_last, $col_first );
        }

        # If the first and last cell are the same write a single cell.
        if ( ( $row_first == $row_last ) && ( $col_first == $col_last ) ) {
            $sqref = $active_cell;
        }
        else {
            $sqref = xl_range( $row_first, $col_first, $row_last, $col_last );
        }

    }
    else {

        # User supplied wrong number or arguments.
        return;
    }

    # Selection isn't set for cell A1.
    return if $sqref eq 'A1';

    $self->{_selections} = [ [ $pane, $active_cell, $sqref ] ];
}


###############################################################################
#
# freeze_panes( $row, $col, $top_row, $left_col )
#
# Set panes and mark them as frozen.
#
sub freeze_panes {

    my $self = shift;

    return unless @_;

    # Check for a cell reference in A1 notation and substitute row and column.
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    my $row      = shift;
    my $col      = shift || 0;
    my $top_row  = shift || $row;
    my $left_col = shift || $col;
    my $type     = shift || 0;

    $self->{_panes} = [ $row, $col, $top_row, $left_col, $type ];
}


###############################################################################
#
# split_panes( $y, $x, $top_row, $left_col )
#
# Set panes and mark them as split.
#
# Implementers note. The API for this method doesn't map well from the XLS
# file format and isn't sufficient to describe all cases of split panes.
# It should probably be something like:
#
#     split_panes( $y, $x, $top_row, $left_col, $offset_row, $offset_col )
#
# I'll look at changing this if it becomes an issue.
#
sub split_panes {

    my $self = shift;

    # Call freeze panes but add the type flag for split panes.
    $self->freeze_panes( @_[ 0 .. 3 ], 2 );
}

# Older method name for backwards compatibility.
*thaw_panes = *split_panes;


###############################################################################
#
# set_portrait()
#
# Set the page orientation as portrait.
#
sub set_portrait {

    my $self = shift;

    $self->{_orientation}        = 1;
    $self->{_page_setup_changed} = 1;
}


###############################################################################
#
# set_landscape()
#
# Set the page orientation as landscape.
#
sub set_landscape {

    my $self = shift;

    $self->{_orientation}        = 0;
    $self->{_page_setup_changed} = 1;
}


###############################################################################
#
# set_page_view()
#
# Set the page view mode for Mac Excel.
#
sub set_page_view {

    my $self = shift;

    $self->{_page_view} = defined $_[0] ? $_[0] : 1;
}


###############################################################################
#
# set_tab_color()
#
# Set the colour of the worksheet tab.
#
sub set_tab_color {

    my $self  = shift;
    my $color = &Excel::Writer::XLSX::Format::_get_color( $_[0] );

    $self->{_tab_color} = $color;
}


###############################################################################
#
# set_paper()
#
# Set the paper type. Ex. 1 = US Letter, 9 = A4
#
sub set_paper {

    my $self       = shift;
    my $paper_size = shift;

    if ( $paper_size ) {
        $self->{_paper_size}         = $paper_size;
        $self->{_page_setup_changed} = 1;
    }
}


###############################################################################
#
# set_header()
#
# Set the page header caption and optional margin.
#
sub set_header {

    my $self = shift;
    my $string = $_[0] || '';

    if ( length $string >= 255 ) {
        carp 'Header string must be less than 255 characters';
        return;
    }

    $self->{_header}                = $string;
    $self->{_margin_header}         = $_[1] || 0.3;
    $self->{_header_footer_changed} = 1;
}


###############################################################################
#
# set_footer()
#
# Set the page footer caption and optional margin.
#
sub set_footer {

    my $self = shift;
    my $string = $_[0] || '';

    if ( length $string >= 255 ) {
        carp 'Footer string must be less than 255 characters';
        return;
    }

    $self->{_footer}                = $string;
    $self->{_margin_footer}         = $_[1] || 0.3;
    $self->{_header_footer_changed} = 1;
}


###############################################################################
#
# center_horizontally()
#
# Center the page horizontally.
#
sub center_horizontally {

    my $self = shift;

    $self->{_print_options_changed} = 1;
    $self->{_hcenter}               = 1;
}


###############################################################################
#
# center_vertically()
#
# Center the page horizontally.
#
sub center_vertically {

    my $self = shift;

    $self->{_print_options_changed} = 1;
    $self->{_vcenter}               = 1;
}


###############################################################################
#
# set_margins()
#
# Set all the page margins to the same value in inches.
#
sub set_margins {

    my $self = shift;

    $self->set_margin_left( $_[0] );
    $self->set_margin_right( $_[0] );
    $self->set_margin_top( $_[0] );
    $self->set_margin_bottom( $_[0] );
}


###############################################################################
#
# set_margins_LR()
#
# Set the left and right margins to the same value in inches.
#
sub set_margins_LR {

    my $self = shift;

    $self->set_margin_left( $_[0] );
    $self->set_margin_right( $_[0] );
}


###############################################################################
#
# set_margins_TB()
#
# Set the top and bottom margins to the same value in inches.
#
sub set_margins_TB {

    my $self = shift;

    $self->set_margin_top( $_[0] );
    $self->set_margin_bottom( $_[0] );
}


###############################################################################
#
# set_margin_left()
#
# Set the left margin in inches.
#
sub set_margin_left {

    my $self    = shift;
    my $margin  = shift;
    my $default = 0.7;

    # Add 0 to ensure the argument is numeric.
    if   ( defined $margin ) { $margin = 0 + $margin }
    else                     { $margin = $default }

    $self->{_margin_left} = $margin;
}


###############################################################################
#
# set_margin_right()
#
# Set the right margin in inches.
#
sub set_margin_right {

    my $self    = shift;
    my $margin  = shift;
    my $default = 0.7;

    # Add 0 to ensure the argument is numeric.
    if   ( defined $margin ) { $margin = 0 + $margin }
    else                     { $margin = $default }

    $self->{_margin_right} = $margin;
}


###############################################################################
#
# set_margin_top()
#
# Set the top margin in inches.
#
sub set_margin_top {

    my $self    = shift;
    my $margin  = shift;
    my $default = 0.75;

    # Add 0 to ensure the argument is numeric.
    if   ( defined $margin ) { $margin = 0 + $margin }
    else                     { $margin = $default }

    $self->{_margin_top} = $margin;
}


###############################################################################
#
# set_margin_bottom()
#
# Set the bottom margin in inches.
#
sub set_margin_bottom {


    my $self    = shift;
    my $margin  = shift;
    my $default = 0.75;

    # Add 0 to ensure the argument is numeric.
    if   ( defined $margin ) { $margin = 0 + $margin }
    else                     { $margin = $default }

    $self->{_margin_bottom} = $margin;
}


###############################################################################
#
# repeat_rows($first_row, $last_row)
#
# Set the rows to repeat at the top of each printed page.
#
sub repeat_rows {

    my $self = shift;

    my $row_min = $_[0];
    my $row_max = $_[1] || $_[0];    # Second row is optional


    # Convert to 1 based.
    $row_min++;
    $row_max++;

    my $area = '$' . $row_min . ':' . '$' . $row_max;

    # Build up the print titles "Sheet1!$1:$2"
    my $sheetname = $self->_quote_sheetname( $self->{_name} );
    $area = $sheetname . "!" . $area;

    $self->{_repeat_rows} = $area;
}


###############################################################################
#
# repeat_columns($first_col, $last_col)
#
# Set the columns to repeat at the left hand side of each printed page. This is
# stored as a <NamedRange> element.
#
sub repeat_columns {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );

        # Returned values $row1 and $row2 aren't required here. Remove them.
        shift @_;    # $row1
        splice @_, 1, 1;    # $row2
    }

    my $col_min = $_[0];
    my $col_max = $_[1] || $_[0];    # Second col is optional

    # Convert to A notation.
    $col_min = xl_col_to_name( $_[0], 1 );
    $col_max = xl_col_to_name( $_[1], 1 );

    my $area = $col_min . ':' . $col_max;

    # Build up the print area range "=Sheet2!C1:C2"
    my $sheetname = $self->_quote_sheetname( $self->{_name} );
    $area = $sheetname . "!" . $area;

    $self->{_repeat_cols} = $area;
}


###############################################################################
#
# print_area($first_row, $first_col, $last_row, $last_col)
#
# Set the print area in the current worksheet. This is stored as a <NamedRange>
# element.
#
sub print_area {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    return if @_ != 4;    # Require 4 parameters

    my ( $row1, $col1, $row2, $col2 ) = @_;

    # Ignore max print area since this is the same as no print area for Excel.
    if (    $row1 == 0
        and $col1 == 0
        and $row2 == $self->{_xls_rowmax} - 1
        and $col2 == $self->{_xls_colmax} - 1 )
    {
        return;
    }

    # Build up the print area range "=Sheet2!R1C1:R2C1"
    my $area = $self->_convert_name_area( $row1, $col1, $row2, $col2 );

    $self->{_print_area} = $area;
}


###############################################################################
#
# autofilter($first_row, $first_col, $last_row, $last_col)
#
# Set the autofilter area in the worksheet.
#
sub autofilter {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    return if @_ != 4;    # Require 4 parameters

    my ( $row1, $col1, $row2, $col2 ) = @_;

    # Reverse max and min values if necessary.
    ( $row1, $row2 ) = ( $row2, $row1 ) if $row2 < $row1;
    ( $col1, $col2 ) = ( $col2, $col1 ) if $col2 < $col1;

    # Build up the print area range "Sheet1!$A$1:$C$13".
    my $area = $self->_convert_name_area( $row1, $col1, $row2, $col2 );
    my $ref = xl_range( $row1, $row2, $col1, $col2 );

    $self->{_autofilter}     = $area;
    $self->{_autofilter_ref} = $ref;
    $self->{_filter_range}   = [ $col1, $col2 ];
}


###############################################################################
#
# filter_column($column, $criteria, ...)
#
# Set the column filter criteria.
#
sub filter_column {

    my $self       = shift;
    my $col        = $_[0];
    my $expression = $_[1];

    croak "Must call autofilter() before filter_column()"
      unless $self->{_autofilter};
    croak "Incorrect number of arguments to filter_column()"
      unless @_ == 2;


    # Check for a column reference in A1 notation and substitute.
    if ( $col =~ /^\D/ ) {
        my $col_letter = $col;

        # Convert col ref to a cell ref and then to a col number.
        ( undef, $col ) = $self->_substitute_cellref( $col . '1' );

        croak "Invalid column '$col_letter'" if $col >= $self->{_xls_colmax};
    }

    my ( $col_first, $col_last ) = @{ $self->{_filter_range} };

    # Reject column if it is outside filter range.
    if ( $col < $col_first or $col > $col_last ) {
        croak "Column '$col' outside autofilter() column range "
          . "($col_first .. $col_last)";
    }


    my @tokens = $self->_extract_filter_tokens( $expression );

    croak "Incorrect number of tokens in expression '$expression'"
      unless ( @tokens == 3 or @tokens == 7 );


    @tokens = $self->_parse_filter_expression( $expression, @tokens );

    # Excel handles single or double custom filters as default filters. We need
    # to check for them and handle them accordingly.
    if ( @tokens == 2 && $tokens[0] == 2 ) {

        # Single equality.
        $self->filter_column_list( $col, $tokens[1] );
    }
    elsif (@tokens == 5
        && $tokens[0] == 2
        && $tokens[2] == 1
        && $tokens[3] == 2 )
    {

        # Double equality with "or" operator.
        $self->filter_column_list( $col, $tokens[1], $tokens[4] );
    }
    else {

        # Non default custom filter.
        $self->{_filter_cols}->{$col} = [@tokens];
        $self->{_filter_type}->{$col} = 0;

    }

    $self->{_filter_on} = 1;
}


###############################################################################
#
# filter_column_list($column, @matches )
#
# Set the column filter criteria in Excel 2007 list style.
#
sub filter_column_list {

    my $self   = shift;
    my $col    = shift;
    my @tokens = @_;

    croak "Must call autofilter() before filter_column_list()"
      unless $self->{_autofilter};
    croak "Incorrect number of arguments to filter_column_list()"
      unless @tokens;

    # Check for a column reference in A1 notation and substitute.
    if ( $col =~ /^\D/ ) {
        my $col_letter = $col;

        # Convert col ref to a cell ref and then to a col number.
        ( undef, $col ) = $self->_substitute_cellref( $col . '1' );

        croak "Invalid column '$col_letter'" if $col >= $self->{_xls_colmax};
    }

    my ( $col_first, $col_last ) = @{ $self->{_filter_range} };

    # Reject column if it is outside filter range.
    if ( $col < $col_first or $col > $col_last ) {
        croak "Column '$col' outside autofilter() column range "
          . "($col_first .. $col_last)";
    }

    $self->{_filter_cols}->{$col} = [@tokens];
    $self->{_filter_type}->{$col} = 1;           # Default style.
    $self->{_filter_on}           = 1;
}


###############################################################################
#
# _extract_filter_tokens($expression)
#
# Extract the tokens from the filter expression. The tokens are mainly non-
# whitespace groups. The only tricky part is to extract string tokens that
# contain whitespace and/or quoted double quotes (Excel's escaped quotes).
#
# Examples: 'x <  2000'
#           'x >  2000 and x <  5000'
#           'x = "foo"'
#           'x = "foo bar"'
#           'x = "foo "" bar"'
#
sub _extract_filter_tokens {

    my $self       = shift;
    my $expression = $_[0];

    return unless $expression;

    my @tokens = ( $expression =~ /"(?:[^"]|"")*"|\S+/g );    #"

    # Remove leading and trailing quotes and unescape other quotes
    for ( @tokens ) {
        s/^"//;                                               #"
        s/"$//;                                               #"
        s/""/"/g;                                             #"
    }

    return @tokens;
}


###############################################################################
#
# _parse_filter_expression(@token)
#
# Converts the tokens of a possibly conditional expression into 1 or 2
# sub expressions for further parsing.
#
# Examples:
#          ('x', '==', 2000) -> exp1
#          ('x', '>',  2000, 'and', 'x', '<', 5000) -> exp1 and exp2
#
sub _parse_filter_expression {

    my $self       = shift;
    my $expression = shift;
    my @tokens     = @_;

    # The number of tokens will be either 3 (for 1 expression)
    # or 7 (for 2  expressions).
    #
    if ( @tokens == 7 ) {

        my $conditional = $tokens[3];

        if ( $conditional =~ /^(and|&&)$/ ) {
            $conditional = 0;
        }
        elsif ( $conditional =~ /^(or|\|\|)$/ ) {
            $conditional = 1;
        }
        else {
            croak "Token '$conditional' is not a valid conditional "
              . "in filter expression '$expression'";
        }

        my @expression_1 =
          $self->_parse_filter_tokens( $expression, @tokens[ 0, 1, 2 ] );
        my @expression_2 =
          $self->_parse_filter_tokens( $expression, @tokens[ 4, 5, 6 ] );

        return ( @expression_1, $conditional, @expression_2 );
    }
    else {
        return $self->_parse_filter_tokens( $expression, @tokens );
    }
}


###############################################################################
#
# _parse_filter_tokens(@token)
#
# Parse the 3 tokens of a filter expression and return the operator and token.
#
sub _parse_filter_tokens {

    my $self       = shift;
    my $expression = shift;
    my @tokens     = @_;

    my %operators = (
        '==' => 2,
        '='  => 2,
        '=~' => 2,
        'eq' => 2,

        '!=' => 5,
        '!~' => 5,
        'ne' => 5,
        '<>' => 5,

        '<'  => 1,
        '<=' => 3,
        '>'  => 4,
        '>=' => 6,
    );

    my $operator = $operators{ $tokens[1] };
    my $token    = $tokens[2];


    # Special handling of "Top" filter expressions.
    if ( $tokens[0] =~ /^top|bottom$/i ) {

        my $value = $tokens[1];

        if (   $value =~ /\D/
            or $value < 1
            or $value > 500 )
        {
            croak "The value '$value' in expression '$expression' "
              . "must be in the range 1 to 500";
        }

        $token = lc $token;

        if ( $token ne 'items' and $token ne '%' ) {
            croak "The type '$token' in expression '$expression' "
              . "must be either 'items' or '%'";
        }

        if ( $tokens[0] =~ /^top$/i ) {
            $operator = 30;
        }
        else {
            $operator = 32;
        }

        if ( $tokens[2] eq '%' ) {
            $operator++;
        }

        $token = $value;
    }


    if ( not $operator and $tokens[0] ) {
        croak "Token '$tokens[1]' is not a valid operator "
          . "in filter expression '$expression'";
    }


    # Special handling for Blanks/NonBlanks.
    if ( $token =~ /^blanks|nonblanks$/i ) {

        # Only allow Equals or NotEqual in this context.
        if ( $operator != 2 and $operator != 5 ) {
            croak "The operator '$tokens[1]' in expression '$expression' "
              . "is not valid in relation to Blanks/NonBlanks'";
        }

        $token = lc $token;

        # The operator should always be 2 (=) to flag a "simple" equality in
        # the binary record. Therefore we convert <> to =.
        if ( $token eq 'blanks' ) {
            if ( $operator == 5 ) {
                $token = ' ';
            }
        }
        else {
            if ( $operator == 5 ) {
                $operator = 2;
                $token    = 'blanks';
            }
            else {
                $operator = 5;
                $token    = ' ';
            }
        }
    }


    # if the string token contains an Excel match character then change the
    # operator type to indicate a non "simple" equality.
    if ( $operator == 2 and $token =~ /[*?]/ ) {
        $operator = 22;
    }


    return ( $operator, $token );
}


###############################################################################
#
# _convert_name_area($first_row, $first_col, $last_row, $last_col)
#
# Convert zero indexed rows and columns to the format required by worksheet
# named ranges, eg, "Sheet1!$A$1:$C$13".
#
sub _convert_name_area {

    my $self = shift;

    my $row_num_1 = $_[0];
    my $col_num_1 = $_[1];
    my $row_num_2 = $_[2];
    my $col_num_2 = $_[3];

    my $range1       = '';
    my $range2       = '';
    my $row_col_only = 0;
    my $area;

    # Convert to A1 notation.
    my $col_char_1 = xl_col_to_name( $col_num_1, 1 );
    my $col_char_2 = xl_col_to_name( $col_num_2, 1 );
    my $row_char_1 = '$' . ( $row_num_1 + 1 );
    my $row_char_2 = '$' . ( $row_num_2 + 1 );

    # We need to handle some special cases that refer to rows or columns only.
    if ( $row_num_1 == 0 and $row_num_2 == $self->{_xls_rowmax} - 1 ) {
        $range1       = $col_char_1;
        $range2       = $col_char_2;
        $row_col_only = 1;
    }
    elsif ( $col_num_1 == 0 and $col_num_2 == $self->{_xls_colmax} - 1 ) {
        $range1       = $row_char_1;
        $range2       = $row_char_2;
        $row_col_only = 1;
    }
    else {
        $range1 = $col_char_1 . $row_char_1;
        $range2 = $col_char_2 . $row_char_2;
    }

    # A repeated range is only written once (if it isn't a special case).
    if ( $range1 eq $range2 && !$row_col_only ) {
        $area = $range1;
    }
    else {
        $area = $range1 . ':' . $range2;
    }

    # Build up the print area range "Sheet1!$A$1:$C$13".
    my $sheetname = $self->_quote_sheetname( $self->{_name} );
    $area = $sheetname . "!" . $area;

    return $area;
}


###############################################################################
#
# hide_gridlines()
#
# Set the option to hide gridlines on the screen and the printed page.
#
# This was mainly useful for Excel 5 where printed gridlines were on by
# default.
#
sub hide_gridlines {

    my $self = shift;
    my $option = defined $_[0] ? $_[0] : 1;    # Default to hiding printed gridlines

    if ( $option == 0 ) {
        $self->{_print_gridlines}       = 1;    # 1 = display, 0 = hide
        $self->{_screen_gridlines}      = 1;
        $self->{_print_options_changed} = 1;
    }
    elsif ( $option == 1 ) {
        $self->{_print_gridlines}  = 0;
        $self->{_screen_gridlines} = 1;
    }
    else {
        $self->{_print_gridlines}  = 0;
        $self->{_screen_gridlines} = 0;
    }
}


###############################################################################
#
# print_row_col_headers()
#
# Set the option to print the row and column headers on the printed page.
# See also the _store_print_headers() method below.
#
sub print_row_col_headers {

    my $self = shift;
    my $headers = defined $_[0] ? $_[0] : 1;

    if ( $headers ) {
        $self->{_print_headers}         = 1;
        $self->{_print_options_changed} = 1;
    }
    else {
        $self->{_print_headers} = 0;
    }
}


###############################################################################
#
# fit_to_pages($width, $height)
#
# Store the vertical and horizontal number of pages that will define the
# maximum area printed.
#
sub fit_to_pages {

    my $self = shift;

    $self->{_fit_page}           = 1;
    $self->{_fit_width}          = defined $_[0] ? $_[0] : 1;
    $self->{_fit_height}         = defined $_[1] ? $_[1] : 1;
    $self->{_page_setup_changed} = 1;
}


###############################################################################
#
# set_h_pagebreaks(@breaks)
#
# Store the horizontal page breaks on a worksheet.
#
sub set_h_pagebreaks {

    my $self = shift;

    push @{ $self->{_hbreaks} }, @_;
}


###############################################################################
#
# set_v_pagebreaks(@breaks)
#
# Store the vertical page breaks on a worksheet.
#
sub set_v_pagebreaks {

    my $self = shift;

    push @{ $self->{_vbreaks} }, @_;
}


###############################################################################
#
# set_zoom( $scale )
#
# Set the worksheet zoom factor.
#
sub set_zoom {

    my $self = shift;
    my $scale = $_[0] || 100;

    # Confine the scale to Excel's range
    if ( $scale < 10 or $scale > 400 ) {
        carp "Zoom factor $scale outside range: 10 <= zoom <= 400";
        $scale = 100;
    }

    $self->{_zoom} = int $scale;
}


###############################################################################
#
# set_print_scale($scale)
#
# Set the scale factor for the printed page.
#
sub set_print_scale {

    my $self = shift;
    my $scale = $_[0] || 100;

    # Confine the scale to Excel's range
    if ( $scale < 10 or $scale > 400 ) {
        carp "Print scale $scale outside range: 10 <= zoom <= 400";
        $scale = 100;
    }

    # Turn off "fit to page" option.
    $self->{_fit_page} = 0;

    $self->{_print_scale}        = int $scale;
    $self->{_page_setup_changed} = 1;
}


###############################################################################
#
# keep_leading_zeros()
#
# Causes the write() method to treat integers with a leading zero as a string.
# This ensures that any leading zeros such, as in zip codes, are maintained.
#
sub keep_leading_zeros {

    my $self = shift;

    if ( defined $_[0] ) {
        $self->{_leading_zeros} = $_[0];
    }
    else {
        $self->{_leading_zeros} = 1;
    }
}


###############################################################################
#
# show_comments()
#
# Make any comments in the worksheet visible.
#
sub show_comments {

    my $self = shift;

    $self->{_comments_visible} = defined $_[0] ? $_[0] : 1;
}


###############################################################################
#
# set_comments_author()
#
# Set the default author of the cell comments.
#
sub set_comments_author {

    my $self = shift;

    $self->{_comments_author} = $_[0] if defined $_[0];
}


###############################################################################
#
# right_to_left()
#
# Display the worksheet right to left for some eastern versions of Excel.
#
sub right_to_left {

    my $self = shift;

    $self->{_right_to_left} = defined $_[0] ? $_[0] : 1;
}


###############################################################################
#
# hide_zero()
#
# Hide cell zero values.
#
sub hide_zero {

    my $self = shift;

    $self->{_show_zeros} = defined $_[0] ? not $_[0] : 0;
}


###############################################################################
#
# print_across()
#
# Set the order in which pages are printed.
#
sub print_across {

    my $self = shift;
    my $page_order = defined $_[0] ? $_[0] : 1;

    if ( $page_order ) {
        $self->{_page_order}         = 1;
        $self->{_page_setup_changed} = 1;
    }
    else {
        $self->{_page_order} = 0;
    }
}


###############################################################################
#
# set_start_page()
#
# Set the start page number.
#
sub set_start_page {

    my $self = shift;
    return unless defined $_[0];

    $self->{_page_start}   = $_[0];
    $self->{_custom_start} = 1;
}


###############################################################################
#
# set_first_row_column()
#
# Set the topmost and leftmost visible row and column.
# TODO: Document this when tested fully for interaction with panes.
#
sub set_first_row_column {

    my $self = shift;

    my $row = $_[0] || 0;
    my $col = $_[1] || 0;

    $row = $self->{_xls_rowmax} if $row > $self->{_xls_rowmax};
    $col = $self->{_xls_colmax} if $col > $self->{_xls_colmax};

    $self->{_first_row} = $row;
    $self->{_first_col} = $col;
}


###############################################################################
#
# add_write_handler($re, $code_ref)
#
# Allow the user to add their own matches and handlers to the write() method.
#
sub add_write_handler {

    my $self = shift;

    return unless @_ == 2;
    return unless ref $_[1] eq 'CODE';

    push @{ $self->{_write_match} }, [@_];
}


###############################################################################
#
# write($row, $col, $token, $format)
#
# Parse $token and call appropriate write method. $row and $column are zero
# indexed. $format is optional.
#
# Returns: return value of called subroutine
#
sub write {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    my $token = $_[2];

    # Handle undefs as blanks
    $token = '' unless defined $token;


    # First try user defined matches.
    for my $aref ( @{ $self->{_write_match} } ) {
        my $re  = $aref->[0];
        my $sub = $aref->[1];

        if ( $token =~ /$re/ ) {
            my $match = &$sub( $self, @_ );
            return $match if defined $match;
        }
    }


    # Match an array ref.
    if ( ref $token eq "ARRAY" ) {
        return $self->write_row( @_ );
    }

    # Match integer with leading zero(s)
    elsif ( $self->{_leading_zeros} and $token =~ /^0\d+$/ ) {
        return $self->write_string( @_ );
    }

    # Match number
    elsif ( $token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/ ) {
        return $self->write_number( @_ );
    }

    # Match http, https or ftp URL
    elsif ( $token =~ m|^[fh]tt?ps?://| ) {
        return $self->write_url( @_ );
    }

    # Match mailto:
    elsif ( $token =~ m/^mailto:/ ) {
        return $self->write_url( @_ );
    }

    # Match internal or external sheet link
    elsif ( $token =~ m[^(?:in|ex)ternal:] ) {
        return $self->write_url( @_ );
    }

    # Match formula
    elsif ( $token =~ /^=/ ) {
        return $self->write_formula( @_ );
    }

    # Match array formula
    elsif ( $token =~ /^{=.*}$/ ) {
        return $self->write_formula( @_ );
    }

    # Match blank
    elsif ( $token eq '' ) {
        splice @_, 2, 1;    # remove the empty string from the parameter list
        return $self->write_blank( @_ );
    }

    # Default: match string
    else {
        return $self->write_string( @_ );
    }
}


###############################################################################
#
# write_row($row, $col, $array_ref, $format)
#
# Write a row of data starting from ($row, $col). Call write_col() if any of
# the elements of the array ref are in turn array refs. This allows the writing
# of 1D or 2D arrays of data in one go.
#
# Returns: the first encountered error value or zero for no errors
#
sub write_row {

    my $self = shift;


    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    # Catch non array refs passed by user.
    if ( ref $_[2] ne 'ARRAY' ) {
        croak "Not an array ref in call to write_row()$!";
    }

    my $row     = shift;
    my $col     = shift;
    my $tokens  = shift;
    my @options = @_;
    my $error   = 0;
    my $ret;

    for my $token ( @$tokens ) {

        # Check for nested arrays
        if ( ref $token eq "ARRAY" ) {
            $ret = $self->write_col( $row, $col, $token, @options );
        }
        else {
            $ret = $self->write( $row, $col, $token, @options );
        }

        # Return only the first error encountered, if any.
        $error ||= $ret;
        $col++;
    }

    return $error;
}


###############################################################################
#
# write_col($row, $col, $array_ref, $format)
#
# Write a column of data starting from ($row, $col). Call write_row() if any of
# the elements of the array ref are in turn array refs. This allows the writing
# of 1D or 2D arrays of data in one go.
#
# Returns: the first encountered error value or zero for no errors
#
sub write_col {

    my $self = shift;


    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    # Catch non array refs passed by user.
    if ( ref $_[2] ne 'ARRAY' ) {
        croak "Not an array ref in call to write_col()$!";
    }

    my $row     = shift;
    my $col     = shift;
    my $tokens  = shift;
    my @options = @_;
    my $error   = 0;
    my $ret;

    for my $token ( @$tokens ) {

        # write() will deal with any nested arrays
        $ret = $self->write( $row, $col, $token, @options );

        # Return only the first error encountered, if any.
        $error ||= $ret;
        $row++;
    }

    return $error;
}


###############################################################################
#
# write_comment($row, $col, $comment)
#
# Write a comment to the specified row and column (zero indexed).
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_comment {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 } # Check the number of args

    my $row = $_[0];
    my $col = $_[1];

    # Check for pairs of optional arguments, i.e. an odd number of args.
    croak "Uneven number of additional arguments" unless @_ % 2;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    $self->{_has_comments} = 1;

    # Process the properties of the cell comment.
    $self->{_comments}->{$row}->{$col} = [ $self->_comment_params(@_) ];
}


###############################################################################
#
# write_number($row, $col, $num, $format)
#
# Write a double to the specified row and column (zero indexed).
# An integer can be written as a double. Excel will display an
# integer. $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_number {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    if ( @_ < 3 ) { return -1 }    # Check the number of args


    my $row  = $_[0];                  # Zero indexed row
    my $col  = $_[1];                  # Zero indexed column
    my $num  = $_[2] + 0;
    my $xf   = $_[3];                  # The cell format
    my $type = 'n';                    # The data type

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    # Write previous row if in in-line string optimization mode.
    if ( $self->{_optimization} == 1 && $row > $self->{_previous_row}) {
        $self->_write_single_row( $row );
    }

    $self->{_table}->[$row]->[$col] = [ $type, $num, $xf ];

    return 0;
}


###############################################################################
#
# write_string ($row, $col, $string, $format)
#
# Write a string to the specified row and column (zero indexed).
# $format is optional.
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 32767 chars
#
sub write_string {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    if ( @_ < 3 ) { return -1 }    # Check the number of args

    my $row  = $_[0];                  # Zero indexed row
    my $col  = $_[1];                  # Zero indexed column
    my $str  = $_[2];
    my $xf   = $_[3];                  # The cell format
    my $type = 's';                    # The data type
    my $index;
    my $str_error = 0;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    # Check that the string is < 32767 chars
    if ( length $str > $self->{_xls_strmax} ) {
        $str = substr( $str, 0, $self->{_xls_strmax} );
        $str_error = -3;
    }

    # Write a shared string or an in-line string based on optimisation level.
    if ( $self->{_optimization} == 0 ) {
        $index = $self->_get_shared_string_index( $str );
    }
    else {
        $index = $str;
    }

    # Write previous row if in in-line string optimization mode.
    if ( $self->{_optimization} == 1 && $row > $self->{_previous_row}) {
        $self->_write_single_row( $row );
    }

    $self->{_table}->[$row]->[$col] = [ $type, $index, $xf ];

    return $str_error;
}


###############################################################################
#
# write_rich_string( $row, $column, $format, $string, ..., $cell_format )
#
# The write_rich_string() method is used to write strings with multiple formats.
# The method receives string fragments prefixed by format objects. The final
# format object is used as the cell format.
#
# Returns  0 : normal termination.
#         -1 : insufficient number of arguments.
#         -2 : row or column out of range.
#         -3 : long string truncated to 32767 chars.
#         -4 : 2 consecutive formats used.
#
sub write_rich_string {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    if ( @_ < 3 ) { return -1 }    # Check the number of args

    my $row    = shift;            # Zero indexed row.
    my $col    = shift;            # Zero indexed column.
    my $str    = '';
    my $xf     = undef;
    my $type   = 's';              # The data type.
    my $length = 0;                # String length.
    my $index;
    my $str_error = 0;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );


    # If the last arg is a format we use it as the cell format.
    if ( ref $_[-1] ) {
        $xf = pop @_;
    }


    # Create a temp XML::Writer object and use it to write the rich string
    # XML to a string.
    open my $str_fh, '>', \$str or die "Failed to open filehandle: $!";

    my $writer = Excel::Writer::XLSX::Package::XMLwriterSimple->new( $str_fh );

    $self->{_rstring} = $writer;

    # Create a temp format with the default font for unformatted fragments.
    my $default = Excel::Writer::XLSX::Format->new();

    # Convert the list of $format, $string tokens to pairs of ($format, $string)
    # except for the first $string fragment which doesn't require a default
    # formatting run. Use the default for strings without a leading format.
    my @fragments;
    my $last = 'format';
    my $pos  = 0;

    for my $token ( @_ ) {
        if ( !ref $token ) {

            # Token is a string.
            if ( $last ne 'format' ) {

                # If previous token wasn't a format add one before the string.
                push @fragments, ( $default, $token );
            }
            else {

                # If previous token was a format just add the string.
                push @fragments, $token;
            }

            $length += length $token;    # Keep track of actual string length.
            $last = 'string';
        }
        else {

            # Can't allow 2 formats in a row.
            if ( $last eq 'format' && $pos > 0 ) {
                return -4;
            }

            # Token is a format object. Add it to the fragment list.
            push @fragments, $token;
            $last = 'format';
        }

        $pos++;
    }


    # If the first token is a string start the <r> element.
    if ( !ref $fragments[0] ) {
        $self->{_rstring}->startTag( 'r' );
    }

    # Write the XML elements for the $format $string fragments.
    for my $token ( @fragments ) {
        if ( ref $token ) {

            # Write the font run.
            $self->{_rstring}->startTag( 'r' );
            $self->_write_font( $token );
        }
        else {

            # Write the string fragment part, with whitespace handling.
            my @attributes = ();

            if ( $token =~ /^\s/ || $token =~ /\s$/ ) {
                push @attributes, ( 'xml:space' => 'preserve' );
            }

            $self->{_rstring}->dataElement( 't', $token, @attributes );
            $self->{_rstring}->endTag( 'r' );
        }
    }

    # Check that the string is < 32767 chars.
    if ( $length > $self->{_xls_strmax} ) {
        return -3;
    }


    # Write a shared string or an in-line string based on optimisation level.
    if ( $self->{_optimization} == 0 ) {
        $index = $self->_get_shared_string_index( $str );
    }
    else {
        $index = $str;
    }

    # Write previous row if in in-line string optimization mode.
    if ( $self->{_optimization} == 1 && $row > $self->{_previous_row}) {
        $self->_write_single_row( $row );
    }

    $self->{_table}->[$row]->[$col] = [ $type, $index, $xf ];

    return 0;
}


###############################################################################
#
# write_blank($row, $col, $format)
#
# Write a blank cell to the specified row and column (zero indexed).
# A blank cell is used to specify formatting without adding a string
# or a number.
#
# A blank cell without a format serves no purpose. Therefore, we don't write
# a BLANK record unless a format is specified. This is mainly an optimisation
# for the write_row() and write_col() methods.
#
# Returns  0 : normal termination (including no format)
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_blank {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    # Check the number of args
    return -1 if @_ < 2;

    # Don't write a blank cell unless it has a format
    return 0 if not defined $_[2];

    my $row  = $_[0];                  # Zero indexed row
    my $col  = $_[1];                  # Zero indexed column
    my $xf   = $_[2];                  # The cell format
    my $type = 'b';                    # The data type

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    # Write previous row if in in-line string optimization mode.
    if ( $self->{_optimization} == 1 && $row > $self->{_previous_row}) {
        $self->_write_single_row( $row );
    }

    $self->{_table}->[$row]->[$col] = [ $type, undef, $xf ];

    return 0;
}


###############################################################################
#
# write_formula($row, $col, $formula, $format)
#
# Write a formula to the specified row and column (zero indexed).
#
# $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_formula {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    if ( @_ < 3 ) { return -1 }    # Check the number of args

    my $row     = $_[0];           # Zero indexed row
    my $col     = $_[1];           # Zero indexed column
    my $formula = $_[2];           # The formula text string
    my $xf      = $_[3];           # The format object.
    my $value   = $_[4];           # Optional formula value.
    my $type    = 'f';             # The data type

    # Hand off array formulas.
    if ( $formula =~ /^{=.*}$/ ) {
        return $self->write_array_formula( $row, $col, $row, $col, $formula,
            $xf, $value );
    }

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    # Remove the = sign if it exists.
    $formula =~ s/^=//;

    # Write previous row if in in-line string optimization mode.
    if ( $self->{_optimization} == 1 && $row > $self->{_previous_row}) {
        $self->_write_single_row( $row );
    }

    $self->{_table}->[$row]->[$col] = [ $type, $formula, $xf, $value ];

    return 0;
}


###############################################################################
#
# write_array_formula($row1, $col1, $row2, $col2, $formula, $format)
#
# Write an array formula to the specified row and column (zero indexed).
#
# $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_array_formula {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    if ( @_ < 5 ) { return -1 }    # Check the number of args

    my $row1    = $_[0];           # First row
    my $col1    = $_[1];           # First column
    my $row2    = $_[2];           # Last row
    my $col2    = $_[3];           # Last column
    my $formula = $_[4];           # The formula text string
    my $xf      = $_[5];           # The format object.
    my $value   = $_[6];           # Optional formula value.
    my $type    = 'a';             # The data type

    # Swap last row/col with first row/col as necessary
    ( $row1, $row2 ) = ( $row2, $row1 ) if $row1 > $row2;
    ( $col1, $col2 ) = ( $col1, $col2 ) if $col1 > $col2;


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row2, $col2 );


    # Define array range
    my $range;

    if ( $row1 == $row2 and $col1 == $col2 ) {
        $range = xl_rowcol_to_cell( $row1, $col1 );

    }
    else {
        $range =
            xl_rowcol_to_cell( $row1, $col1 ) . ':'
          . xl_rowcol_to_cell( $row2, $col2 );
    }

    # Remove array formula braces and the leading =.
    $formula =~ s/^{(.*)}$/$1/;
    $formula =~ s/^=//;

    # Write previous row if in in-line string optimization mode.
    my $row = $row1;
    if ( $self->{_optimization} == 1 && $row > $self->{_previous_row}) {
        $self->_write_single_row( $row );
    }

    $self->{_table}->[$row1]->[$col1] =
      [ $type, $formula, $xf, $range, $value ];

    return 0;
}


###############################################################################
#
# outline_settings($visible, $symbols_below, $symbols_right, $auto_style)
#
# This method sets the properties for outlining and grouping. The defaults
# correspond to Excel's defaults.
#
sub outline_settings {

    my $self = shift;

    $self->{_outline_on}    = defined $_[0] ? $_[0] : 1;
    $self->{_outline_below} = defined $_[1] ? $_[1] : 1;
    $self->{_outline_right} = defined $_[2] ? $_[2] : 1;
    $self->{_outline_style} = $_[3] || 0;

    $self->{_outline_changed} = 1;
}


###############################################################################
#
# write_url($row, $col, $url, $string, $format)
#
# Write a hyperlink. This is comprised of two elements: the visible label and
# the invisible link. The visible label is the same as the link unless an
# alternative string is specified. The label is written using the
# write_string() method. Therefore the max characters string limit applies.
# $string and $format are optional and their order is interchangeable.
#
# The hyperlink can be to a http, ftp, mail, internal sheet, or external
# directory url.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 32767 chars
#
sub write_url {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    if ( @_ < 3 ) { return -1 }    # Check the number of args


    # Reverse the order of $string and $format if necessary. We work on a copy
    # in order to protect the callers args. We don't use "local @_" in case of
    # perl50005 threads.
    my @args = @_;
    ( $args[3], $args[4] ) = ( $args[4], $args[3] ) if ref $args[3];


    my $row       = $args[0];                  # Zero indexed row
    my $col       = $args[1];                  # Zero indexed column
    my $url       = $args[2];                  # URL string
    my $str       = $args[3];                  # Alternative label
    my $xf        = $args[4];                  # Cell format
    my $tip       = $args[5];                  # Tool tip
    my $type      = 'l';                       # XML data type
    my $link_type = 1;


    # Remove the URI scheme from internal links.
    if ( $url =~ s/^internal:// ) {
        $link_type = 2;
    }

    # Remove the URI scheme from external links.
    if ( $url =~ s/^external:// ) {
        $link_type = 3;
    }

    # The displayed string defaults to the url string.
    $str = $url unless defined $str;

    # For external links change the directory separator from Unix to Dos.
    if ( $link_type == 3 ) {
        $url =~ s[/][\\]g;
        $str =~ s[/][\\]g;
    }

    # Strip the mailto header.
    $str =~ s/^mailto://;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    # Check that the string is < 32767 chars
    my $str_error = 0;
    if ( length $str > $self->{_xls_strmax} ) {
        $str = substr( $str, 0, $self->{_xls_strmax} );
        $str_error = -3;
    }

    # Store the URL displayed text in the shared string table.
    my $index = $self->_get_shared_string_index( $str );

    # External links to URLs and to other Excel workbooks have slightly
    # different characteristics that we have to account for.
    if ( $link_type == 1 ) {

        # Ordinary URL style external links don't have a "location" string.
        $str = undef;
    }
    elsif ( $link_type == 3 ) {

        # External Workbook links need to be modified into the right format.
        # The URL will look something like 'c:\temp\file.xlsx#Sheet!A1'.
        # We need the part to the left of the # as the URL and the part to
        # the right as the "location" string (if it exists)
        ( $url, $str ) = split /#/, $url;

        # Add the file:/// URI to the $url if non-local.
        if ( $url =~ m{[\\/]} && $url !~ m{^\.\.} ) {
            $url = 'file:///' . $url;
        }

        # Treat as a default external link now that the data has been modified.
        $link_type = 1;
    }

    # Write previous row if in in-line string optimization mode.
    if ( $self->{_optimization} == 1 && $row > $self->{_previous_row}) {
        $self->_write_single_row( $row );
    }

    $self->{_table}->[$row]->[$col] =

      # 0      1       2    3           4     5     6
      [ $type, $index, $xf, $link_type, $url, $str, $tip ];

    return $str_error;
}


###############################################################################
#
# write_date_time ($row, $col, $string, $format)
#
# Write a datetime string in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format as a
# number representing an Excel date. $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : Invalid date_time, written as string
#
sub write_date_time {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    if ( @_ < 3 ) { return -1 }    # Check the number of args

    my $row  = $_[0];                  # Zero indexed row
    my $col  = $_[1];                  # Zero indexed column
    my $str  = $_[2];
    my $xf   = $_[3];                  # The cell format
    my $type = 'n';                    # The data type


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    my $str_error = 0;
    my $date_time = $self->convert_date_time( $str );

    # If the date isn't valid then write it as a string.
    if ( !defined $date_time ) {
        return $self->write_string( @_ );
    }

    # Write previous row if in in-line string optimization mode.
    if ( $self->{_optimization} == 1 && $row > $self->{_previous_row}) {
        $self->_write_single_row( $row );
    }

    $self->{_table}->[$row]->[$col] = [ $type, $date_time, $xf ];

    return $str_error;
}


###############################################################################
#
# convert_date_time($date_time_string)
#
# The function takes a date and time in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format
# and converts it to a decimal number representing a valid Excel date.
#
# Dates and times in Excel are represented by real numbers. The integer part of
# the number stores the number of days since the epoch and the fractional part
# stores the percentage of the day in seconds. The epoch can be either 1900 or
# 1904.
#
# Parameter: Date and time string in one of the following formats:
#               yyyy-mm-ddThh:mm:ss.ss  # Standard
#               yyyy-mm-ddT             # Date only
#                         Thh:mm:ss.ss  # Time only
#
# Returns:
#            A decimal number representing a valid Excel date, or
#            undef if the date is invalid.
#
sub convert_date_time {

    my $self      = shift;
    my $date_time = $_[0];

    my $days    = 0;    # Number of days since epoch
    my $seconds = 0;    # Time expressed as fraction of 24h hours in seconds

    my ( $year, $month, $day );
    my ( $hour, $min,   $sec );


    # Strip leading and trailing whitespace.
    $date_time =~ s/^\s+//;
    $date_time =~ s/\s+$//;

    # Check for invalid date char.
    return if $date_time =~ /[^0-9T:\-\.Z]/;

    # Check for "T" after date or before time.
    return unless $date_time =~ /\dT|T\d/;

    # Strip trailing Z in ISO8601 date.
    $date_time =~ s/Z$//;


    # Split into date and time.
    my ( $date, $time ) = split /T/, $date_time;


    # We allow the time portion of the input DateTime to be optional.
    if ( $time ne '' ) {

        # Match hh:mm:ss.sss+ where the seconds are optional
        if ( $time =~ /^(\d\d):(\d\d)(:(\d\d(\.\d+)?))?/ ) {
            $hour = $1;
            $min  = $2;
            $sec  = $4 || 0;
        }
        else {
            return undef;    # Not a valid time format.
        }

        # Some boundary checks
        return if $hour >= 24;
        return if $min >= 60;
        return if $sec >= 60;

        # Excel expresses seconds as a fraction of the number in 24 hours.
        $seconds = ( $hour * 60 * 60 + $min * 60 + $sec ) / ( 24 * 60 * 60 );
    }


    # We allow the date portion of the input DateTime to be optional.
    return $seconds if $date eq '';


    # Match date as yyyy-mm-dd.
    if ( $date =~ /^(\d\d\d\d)-(\d\d)-(\d\d)$/ ) {
        $year  = $1;
        $month = $2;
        $day   = $3;
    }
    else {
        return undef;    # Not a valid date format.
    }

    # Set the epoch as 1900 or 1904. Defaults to 1900.
    my $date_1904 = $self->{_1904};


    # Special cases for Excel.
    if ( not $date_1904 ) {
        return $seconds      if $date eq '1899-12-31';    # Excel 1900 epoch
        return $seconds      if $date eq '1900-01-00';    # Excel 1900 epoch
        return 60 + $seconds if $date eq '1900-02-29';    # Excel false leapday
    }


    # We calculate the date by calculating the number of days since the epoch
    # and adjust for the number of leap days. We calculate the number of leap
    # days by normalising the year in relation to the epoch. Thus the year 2000
    # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
    #
    my $epoch  = $date_1904 ? 1904 : 1900;
    my $offset = $date_1904 ? 4    : 0;
    my $norm   = 300;
    my $range  = $year - $epoch;


    # Set month days and check for leap year.
    my @mdays = ( 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );
    my $leap = 0;
    $leap = 1 if $year % 4 == 0 and $year % 100 or $year % 400 == 0;
    $mdays[1] = 29 if $leap;


    # Some boundary checks
    return if $year < $epoch or $year > 9999;
    return if $month < 1     or $month > 12;
    return if $day < 1       or $day > $mdays[ $month - 1 ];

    # Accumulate the number of days since the epoch.
    $days = $day;    # Add days for current month
    $days += $mdays[$_] for 0 .. $month - 2;    # Add days for past months
    $days += $range * 365;                      # Add days for past years
    $days += int( ( $range ) / 4 );             # Add leapdays
    $days -= int( ( $range + $offset ) / 100 ); # Subtract 100 year leapdays
    $days += int( ( $range + $offset + $norm ) / 400 );  # Add 400 year leapdays
    $days -= $leap;                                      # Already counted above


    # Adjust for Excel erroneously treating 1900 as a leap year.
    $days++ if $date_1904 == 0 and $days > 59;

    return $days + $seconds;
}


###############################################################################
#
# set_row($row, $height, $XF, $hidden, $level, $collapsed)
#
# This method is used to set the height and XF format for a row.
#
sub set_row {

    my $self      = shift;
    my $row       = shift;          # Row Number.
    my $height    = shift;          # Row height.
    my $xf        = shift;          # Format object.
    my $hidden    = shift || 0;     # Hidden flag.
    my $level     = shift || 0;     # Outline level.
    my $collapsed = shift || 0;     # Collapsed row.

    return unless defined $row;     # Ensure at least $row is specified.

    # Check that row and col are valid and store max and min values.
    return -2 if $self->_check_dimensions( $row, 0 );

    $height = 15 if !defined $height;

    # If the height is 0 the row is hidden and the height is the default.
    if ( $height == 0 ) {
        $hidden = 1;
        $height = 15;
    }

    # Set the limits for the outline levels (0 <= x <= 7).
    $level = 0 if $level < 0;
    $level = 7 if $level > 7;

    if ( $level > $self->{_outline_row_level} ) {
        $self->{_outline_row_level} = $level;
    }

    # Store the row properties.
    $self->{_set_rows}->{$row} = [ $height, $xf, $hidden, $level, $collapsed ];

    # Store the row change to allow optimisations.
    $self->{_row_size_changed} = 1;

    # Store the row sizes for use when calculating image vertices.
    $self->{_row_sizes}->{$row} = $height;
}


###############################################################################
#
# merge_range($first_row, $first_col, $last_row, $last_col, $string, $format)
#
# Merge a range of cells. The first cell should contain the data and the others
# should be blank. All cells should contain the same format.
#
sub merge_range {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }
    croak "Incorrect number of arguments" if @_ < 6;
    croak "Fifth parameter must be a format object" unless ref $_[5];

    my $row_first  = shift;
    my $col_first  = shift;
    my $row_last   = shift;
    my $col_last   = shift;
    my $string     = shift;
    my $format     = shift;
    my @extra_args = @_;      # For write_url().

    # Excel doesn't allow a single cell to be merged
    if ( $row_first == $row_last and $col_first == $col_last ) {
        croak "Can't merge single cell";
    }

    # Swap last row/col with first row/col as necessary
    ( $row_first, $row_last ) = ( $row_last, $row_first )
      if $row_first > $row_last;
    ( $col_first, $col_last ) = ( $col_last, $col_first )
      if $col_first > $col_last;

    # Check that column number is valid and store the max value
    return if $self->_check_dimensions( $row_last, $col_last );

    # Store the merge range.
    push @{ $self->{_merge} }, [ $row_first, $col_first, $row_last, $col_last ];

    # Write the first cell
    $self->write( $row_first, $col_first, $string, $format, @extra_args );

    # Pad out the rest of the area with formatted blank cells.
    for my $row ( $row_first .. $row_last ) {
        for my $col ( $col_first .. $col_last ) {
            next if $row == $row_first and $col == $col_first;
            $self->write_blank( $row, $col, $format );
        }
    }
}


###############################################################################
#
# merge_range_type()
#
# Same as merge_range() above except the type of write() is specified.
#
sub merge_range_type {

    my $self = shift;
    my $type = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    my $row_first = shift;
    my $col_first = shift;
    my $row_last  = shift;
    my $col_last  = shift;
    my $format;

    # Get the format. It can be in different positions for the different types.
    if (   $type eq 'array_formula'
        || $type eq 'blank'
        || $type eq 'rich_string' )
    {

        # The format is the last element.
        $format = $_[-1];
    }
    else {

        # Or else it is after the token.
        $format = $_[1];
    }

    # Check that there is a format object.
    croak "Format object missing or in an incorrect position" unless ref $format;

    # Excel doesn't allow a single cell to be merged
    if ( $row_first == $row_last and $col_first == $col_last ) {
        croak "Can't merge single cell";
    }

    # Swap last row/col with first row/col as necessary
    ( $row_first, $row_last ) = ( $row_last, $row_first )
      if $row_first > $row_last;
    ( $col_first, $col_last ) = ( $col_last, $col_first )
      if $col_first > $col_last;

    # Check that column number is valid and store the max value
    return if $self->_check_dimensions( $row_last, $col_last );

    # Store the merge range.
    push @{ $self->{_merge} }, [ $row_first, $col_first, $row_last, $col_last ];

    # Write the first cell
    if ( $type eq 'string' ) {
        $self->write_string( $row_first, $col_first, @_ );
    }
    elsif ( $type eq 'number' ) {
        $self->write_number( $row_first, $col_first, @_ );
    }
    elsif ( $type eq 'blank' ) {
        $self->write_blank( $row_first, $col_first, @_ );
    }
    elsif ( $type eq 'date_time' ) {
        $self->write_date_time( $row_first, $col_first, @_ );
    }
    elsif ( $type eq 'rich_string' ) {
        $self->write_rich_string( $row_first, $col_first, @_ );
    }
    elsif ( $type eq 'url' ) {
        $self->write_url( $row_first, $col_first, @_ );
    }
    elsif ( $type eq 'formula' ) {
        $self->write_formula( $row_first, $col_first, @_ );
    }
    elsif ( $type eq 'array_formula' ) {
        $self->write_formula_array( $row_first, $col_first, @_ );
    }
    else {
        croak "Unknown type '$type'";
    }

    # Pad out the rest of the area with formatted blank cells.
    for my $row ( $row_first .. $row_last ) {
        for my $col ( $col_first .. $col_last ) {
            next if $row == $row_first and $col == $col_first;
            $self->write_blank( $row, $col, $format );
        }
    }
}


###############################################################################
#
# data_validation($row, $col, {...})
#
# This method handles the interface to Excel data validation.
# Somewhat ironically this requires a lot of validation code since the
# interface is flexible and covers a several types of data validation.
#
# We allow data validation to be called on one cell or a range of cells. The
# hashref contains the validation parameters and must be the last param:
#    data_validation($row, $col, {...})
#    data_validation($first_row, $first_col, $last_row, $last_col, {...})
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : incorrect parameter.
#
sub data_validation {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    # Check for a valid number of args.
    if ( @_ != 5 && @_ != 3 ) { return -1 }

    # The final hashref contains the validation parameters.
    my $param = pop;

    # Make the last row/col the same as the first if not defined.
    my ( $row1, $col1, $row2, $col2 ) = @_;
    if ( !defined $row2 ) {
        $row2 = $row1;
        $col2 = $col1;
    }

    # Check that row and col are valid without storing the values.
    return -2 if $self->_check_dimensions( $row1, $col1, 1, 1 );
    return -2 if $self->_check_dimensions( $row2, $col2, 1, 1 );


    # Check that the last parameter is a hash list.
    if ( ref $param ne 'HASH' ) {
        carp "Last parameter '$param' in data_validation() must be a hash ref";
        return -3;
    }

    # List of valid input parameters.
    my %valid_parameter = (
        validate      => 1,
        criteria      => 1,
        value         => 1,
        source        => 1,
        minimum       => 1,
        maximum       => 1,
        ignore_blank  => 1,
        dropdown      => 1,
        show_input    => 1,
        input_title   => 1,
        input_message => 1,
        show_error    => 1,
        error_title   => 1,
        error_message => 1,
        error_type    => 1,
        other_cells   => 1,
    );

    # Check for valid input parameters.
    for my $param_key ( keys %$param ) {
        if ( not exists $valid_parameter{$param_key} ) {
            carp "Unknown parameter '$param_key' in data_validation()";
            return -3;
        }
    }

    # Map alternative parameter names 'source' or 'minimum' to 'value'.
    $param->{value} = $param->{source}  if defined $param->{source};
    $param->{value} = $param->{minimum} if defined $param->{minimum};

    # 'validate' is a required parameter.
    if ( not exists $param->{validate} ) {
        carp "Parameter 'validate' is required in data_validation()";
        return -3;
    }


    # List of  valid validation types.
    my %valid_type = (
        'any'          => 'none',
        'any value'    => 'none',
        'whole number' => 'whole',
        'whole'        => 'whole',
        'integer'      => 'whole',
        'decimal'      => 'decimal',
        'list'         => 'list',
        'date'         => 'date',
        'time'         => 'time',
        'text length'  => 'textLength',
        'length'       => 'textLength',
        'custom'       => 'custom',
    );


    # Check for valid validation types.
    if ( not exists $valid_type{ lc( $param->{validate} ) } ) {
        carp "Unknown validation type '$param->{validate}' for parameter "
          . "'validate' in data_validation()";
        return -3;
    }
    else {
        $param->{validate} = $valid_type{ lc( $param->{validate} ) };
    }


    # No action is required for validation type 'any'.
    # TODO: we should perhaps store 'any' for message only validations.
    return 0 if $param->{validate} eq 'none';


    # The list and custom validations don't have a criteria so we use a default
    # of 'between'.
    if ( $param->{validate} eq 'list' || $param->{validate} eq 'custom' ) {
        $param->{criteria} = 'between';
        $param->{maximum}  = undef;
    }

    # 'criteria' is a required parameter.
    if ( not exists $param->{criteria} ) {
        carp "Parameter 'criteria' is required in data_validation()";
        return -3;
    }


    # List of valid criteria types.
    my %criteria_type = (
        'between'                  => 'between',
        'not between'              => 'notBetween',
        'equal to'                 => 'equal',
        '='                        => 'equal',
        '=='                       => 'equal',
        'not equal to'             => 'notEqual',
        '!='                       => 'notEqual',
        '<>'                       => 'notEqual',
        'greater than'             => 'greaterThan',
        '>'                        => 'greaterThan',
        'less than'                => 'lessThan',
        '<'                        => 'lessThan',
        'greater than or equal to' => 'greaterThanOrEqual',
        '>='                       => 'greaterThanOrEqual',
        'less than or equal to'    => 'lessThanOrEqual',
        '<='                       => 'lessThanOrEqual',
    );

    # Check for valid criteria types.
    if ( not exists $criteria_type{ lc( $param->{criteria} ) } ) {
        carp "Unknown criteria type '$param->{criteria}' for parameter "
          . "'criteria' in data_validation()";
        return -3;
    }
    else {
        $param->{criteria} = $criteria_type{ lc( $param->{criteria} ) };
    }


    # 'Between' and 'Not between' criteria require 2 values.
    if ( $param->{criteria} eq 'between' || $param->{criteria} eq 'notBetween' )
    {
        if ( not exists $param->{maximum} ) {
            carp "Parameter 'maximum' is required in data_validation() "
              . "when using 'between' or 'not between' criteria";
            return -3;
        }
    }
    else {
        $param->{maximum} = undef;
    }


    # List of valid error dialog types.
    my %error_type = (
        'stop'        => 0,
        'warning'     => 1,
        'information' => 2,
    );

    # Check for valid error dialog types.
    if ( not exists $param->{error_type} ) {
        $param->{error_type} = 0;
    }
    elsif ( not exists $error_type{ lc( $param->{error_type} ) } ) {
        carp "Unknown criteria type '$param->{error_type}' for parameter "
          . "'error_type' in data_validation()";
        return -3;
    }
    else {
        $param->{error_type} = $error_type{ lc( $param->{error_type} ) };
    }


    # Convert date/times value if required.
    if ( $param->{validate} eq 'date' || $param->{validate} eq 'time' ) {
        if ( $param->{value} =~ /T/ ) {
            my $date_time = $self->convert_date_time( $param->{value} );

            if ( !defined $date_time ) {
                carp "Invalid date/time value '$param->{value}' "
                  . "in data_validation()";
                return -3;
            }
            else {
                $param->{value} = $date_time;
            }
        }
        if ( defined $param->{maximum} && $param->{maximum} =~ /T/ ) {
            my $date_time = $self->convert_date_time( $param->{maximum} );

            if ( !defined $date_time ) {
                carp "Invalid date/time value '$param->{maximum}' "
                  . "in data_validation()";
                return -3;
            }
            else {
                $param->{maximum} = $date_time;
            }
        }
    }


    # Set some defaults if they haven't been defined by the user.
    $param->{ignore_blank} = 1 if !defined $param->{ignore_blank};
    $param->{dropdown}     = 1 if !defined $param->{dropdown};
    $param->{show_input}   = 1 if !defined $param->{show_input};
    $param->{show_error}   = 1 if !defined $param->{show_error};


    # These are the cells to which the validation is applied.
    $param->{cells} = [ [ $row1, $col1, $row2, $col2 ] ];

    # A (for now) undocumented parameter to pass additional cell ranges.
    if ( exists $param->{other_cells} ) {

        push @{ $param->{cells} }, @{ $param->{other_cells} };
    }

    # Store the validation information until we close the worksheet.
    push @{ $self->{_validations} }, $param;
}


###############################################################################
#
# conditional_formatting($row, $col, {...})
#
# This method handles the interface to Excel conditional formatting.
#
# We allow the format to be called on one cell or a range of cells. The
# hashref contains the formatting parameters and must be the last param:
#    conditional_formatting($row, $col, {...})
#    conditional_formatting($first_row, $first_col, $last_row, $last_col, {...})
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : incorrect parameter.
#
sub conditional_formatting {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    # Check for a valid number of args.
    if ( @_ != 5 && @_ != 3 ) { return -1 }

    # The final hashref contains the validation parameters.
    my $param = pop;

    # Make the last row/col the same as the first if not defined.
    my ( $row1, $col1, $row2, $col2 ) = @_;
    if ( !defined $row2 ) {
        $row2 = $row1;
        $col2 = $col1;
    }

    # Check that row and col are valid without storing the values.
    return -2 if $self->_check_dimensions( $row1, $col1, 1, 1 );
    return -2 if $self->_check_dimensions( $row2, $col2, 1, 1 );


    # Check that the last parameter is a hash list.
    if ( ref $param ne 'HASH' ) {
        carp "Last parameter '$param' in conditional_formatting() "
          . "must be a hash ref";
        return -3;
    }

    # List of valid input parameters.
    my %valid_parameter = (
        type     => 1,
        format   => 1,
        criteria => 1,
        value    => 1,
        minimum  => 1,
        maximum  => 1,
    );

    # Check for valid input parameters.
    for my $param_key ( keys %$param ) {
        if ( not exists $valid_parameter{$param_key} ) {
            carp "Unknown parameter '$param_key' in conditional_formatting()";
            return -3;
        }
    }

    # 'type' is a required parameter.
    if ( not exists $param->{type} ) {
        carp "Parameter 'type' is required in conditional_formatting()";
        return -3;
    }


    # List of  valid validation types.
    my %valid_type = (
        'cell'          => 'cellIs',
        'date'          => 'date',
        'time'          => 'time',
        'average'       => 'aboveAverage',
        'duplicate'     => 'duplicateValues',
        'unique'        => 'uniqueValues',
        'top'           => 'top10',
        'bottom'        => 'top10',
        'text'          => 'text',
        'time_period'   => 'timePeriod',
        'blanks'        => 'containsBlanks',
        'no_blanks'     => 'notContainsBlanks',
        'errors'        => 'containsErrors',
        'no_errors'     => 'notContainsErrors',
        '2_color_scale' => '2_color_scale',
        '3_color_scale' => '3_color_scale',
        'data_bar'      => 'dataBar',
        'formula'       => 'expression',
    );


    # Check for valid validation types.
    if ( not exists $valid_type{ lc( $param->{type} ) } ) {
        carp "Unknown validation type '$param->{type}' for parameter "
          . "'type' in conditional_formatting()";
        return -3;
    }
    else {
        $param->{direction} = 'bottom' if $param->{type} eq 'bottom';
        $param->{type} = $valid_type{ lc( $param->{type} ) };
    }


    # List of valid criteria types.
    my %criteria_type = (
        'between'                  => 'between',
        'not between'              => 'notBetween',
        'equal to'                 => 'equal',
        '='                        => 'equal',
        '=='                       => 'equal',
        'not equal to'             => 'notEqual',
        '!='                       => 'notEqual',
        '<>'                       => 'notEqual',
        'greater than'             => 'greaterThan',
        '>'                        => 'greaterThan',
        'less than'                => 'lessThan',
        '<'                        => 'lessThan',
        'greater than or equal to' => 'greaterThanOrEqual',
        '>='                       => 'greaterThanOrEqual',
        'less than or equal to'    => 'lessThanOrEqual',
        '<='                       => 'lessThanOrEqual',
        'containing'               => 'containsText',
        'not containing'           => 'notContains',
        'begins with'              => 'beginsWith',
        'ends with'                => 'endsWith',
        'yesterday'                => 'yesterday',
        'today'                    => 'today',
        'last 7 days'              => 'last7Days',
        'last week'                => 'lastWeek',
        'this week'                => 'thisWeek',
        'next week'                => 'nextWeek',
        'last month'               => 'lastMonth',
        'this month'               => 'thisMonth',
        'next month'               => 'nextMonth',
    );

    # Check for valid criteria types.
    if ( exists $criteria_type{ lc( $param->{criteria} ) } ) {
        $param->{criteria} = $criteria_type{ lc( $param->{criteria} ) };
    }

    # Convert date/times value if required.
    if ( $param->{type} eq 'date' || $param->{type} eq 'time' ) {
        $param->{type} = 'cellIs';

        if ( defined $param->{value} && $param->{value} =~ /T/ ) {
            my $date_time = $self->convert_date_time( $param->{value} );

            if ( !defined $date_time ) {
                carp "Invalid date/time value '$param->{value}' "
                  . "in conditional_formatting()";
                return -3;
            }
            else {
                $param->{value} = $date_time;
            }
        }

        if ( defined $param->{minimum} && $param->{minimum} =~ /T/ ) {
            my $date_time = $self->convert_date_time( $param->{minimum} );

            if ( !defined $date_time ) {
                carp "Invalid date/time value '$param->{minimum}' "
                  . "in conditional_formatting()";
                return -3;
            }
            else {
                $param->{minimum} = $date_time;
            }
        }

        if ( defined $param->{maximum} && $param->{maximum} =~ /T/ ) {
            my $date_time = $self->convert_date_time( $param->{maximum} );

            if ( !defined $date_time ) {
                carp "Invalid date/time value '$param->{maximum}' "
                  . "in conditional_formatting()";
                return -3;
            }
            else {
                $param->{maximum} = $date_time;
            }
        }
    }

    # Set the formatting range.
    my $range      = '';
    my $start_cell = '';    # Use for formulas.

    # Swap last row/col for first row/col as necessary
    if ( $row1 > $row2 ) {
        ( $row1, $row2 ) = ( $row2, $row1 );
    }

    if ( $col1 > $col2 ) {
        ( $col1, $col2 ) = ( $col2, $col1 );
    }

    # If the first and last cell are the same write a single cell.
    if ( ( $row1 == $row2 ) && ( $col1 == $col2 ) ) {
        $range = xl_rowcol_to_cell( $row1, $col1 );
        $start_cell = $range;
    }
    else {
        $range = xl_range( $row1, $row2, $col1, $col2 );
        $start_cell = xl_rowcol_to_cell( $row1, $col1 );
    }

    # Get the dxf format index.
    if ( defined $param->{format} && ref $param->{format} ) {
        $param->{format} = $param->{format}->get_dxf_index();
    }

    # Set the priority based on the order of adding.
    $param->{priority} = $self->{_dxf_priority}++;

    # Special handling of text criteria.
    if ( $param->{type} eq 'text' ) {

        if ( $param->{criteria} eq 'containsText' ) {
            $param->{type}     = 'containsText';
            $param->{formula}  = sprintf 'NOT(ISERROR(SEARCH("%s",%s)))',
              $param->{value}, $start_cell;
        }
        elsif ( $param->{criteria} eq 'notContains' ) {
            $param->{type}     = 'notContainsText';
            $param->{formula}  = sprintf 'ISERROR(SEARCH("%s",%s))',
              $param->{value}, $start_cell;
        }
        elsif ( $param->{criteria} eq 'beginsWith' ) {
            $param->{type}     = 'beginsWith';
            $param->{formula}  = sprintf 'LEFT(%s,1)="%s"',
              $start_cell, $param->{value};
        }
        elsif ( $param->{criteria} eq 'endsWith' ) {
            $param->{type}     = 'endsWith';
            $param->{formula}  = sprintf 'RIGHT(%s,1)="%s"',
              $start_cell, $param->{value};
        }
        else {
            carp "Invalid text criteria '$param->{criteria}' "
              . "in conditional_formatting()";
        }
    }

    # Special handling of time time_period criteria.
    if ( $param->{type} eq 'timePeriod' ) {

        if ( $param->{criteria} eq 'yesterday' ) {
            $param->{formula} = sprintf 'FLOOR(%s,1)=TODAY()-1', $start_cell;
        }
        elsif ( $param->{criteria} eq 'today' ) {
            $param->{formula} = sprintf 'FLOOR(%s,1)=TODAY()', $start_cell;
        }
        elsif ( $param->{criteria} eq 'tomorrow' ) {
            $param->{formula} = sprintf 'FLOOR(%s,1)=TODAY()+1', $start_cell;
        }
        elsif ( $param->{criteria} eq 'last7Days' ) {
            $param->{formula} =
              sprintf 'AND(TODAY()-FLOOR(%s,1)<=6,FLOOR(%s,1)<=TODAY())',
              $start_cell, $start_cell;
        }
        elsif ( $param->{criteria} eq 'lastWeek' ) {
            $param->{formula} =
              sprintf 'AND(TODAY()-ROUNDDOWN(%s,0)>=(WEEKDAY(TODAY())),'
              . 'TODAY()-ROUNDDOWN(%s,0)<(WEEKDAY(TODAY())+7))',
              $start_cell, $start_cell;
        }
        elsif ( $param->{criteria} eq 'thisWeek' ) {
            $param->{formula} =
              sprintf 'AND(TODAY()-ROUNDDOWN(%s,0)<=WEEKDAY(TODAY())-1,'
              . 'ROUNDDOWN(%s,0)-TODAY()<=7-WEEKDAY(TODAY()))',
              $start_cell, $start_cell;
        }
        elsif ( $param->{criteria} eq 'nextWeek' ) {
            $param->{formula} =
              sprintf 'AND(ROUNDDOWN(%s,0)-TODAY()>(7-WEEKDAY(TODAY())),'
              . 'ROUNDDOWN(%s,0)-TODAY()<(15-WEEKDAY(TODAY())))',
              $start_cell, $start_cell;
        }
        elsif ( $param->{criteria} eq 'lastMonth' ) {
            $param->{formula} =
              sprintf
              'AND(MONTH(%s)=MONTH(TODAY())-1,OR(YEAR(%s)=YEAR(TODAY()),'
              . 'AND(MONTH(%s)=1,YEAR(A1)=YEAR(TODAY())-1)))',
              $start_cell, $start_cell, $start_cell;
        }
        elsif ( $param->{criteria} eq 'thisMonth' ) {
            $param->{formula} =
              sprintf 'AND(MONTH(%s)=MONTH(TODAY()),YEAR(%s)=YEAR(TODAY()))',
              $start_cell, $start_cell, $start_cell;
        }
        elsif ( $param->{criteria} eq 'nextMonth' ) {
            $param->{formula} =
              sprintf
              'AND(MONTH(%s)=MONTH(TODAY())+1,OR(YEAR(%s)=YEAR(TODAY()),'
              . 'AND(MONTH(%s)=12,YEAR(%s)=YEAR(TODAY())+1)))',
              $start_cell, $start_cell, $start_cell, $start_cell;
        }
        else {
            carp "Invalid time_period criteria '$param->{criteria}' "
              . "in conditional_formatting()";
        }
    }


    # Special handling of blanks/error types.
    if ( $param->{type} eq 'containsBlanks' ) {
        $param->{formula} = sprintf 'LEN(TRIM(%s))=0', $start_cell;
    }

    if ( $param->{type} eq 'notContainsBlanks' ) {
        $param->{formula} = sprintf 'LEN(TRIM(%s))>0', $start_cell;
    }

    if ( $param->{type} eq 'containsErrors' ) {
        $param->{formula} = sprintf 'ISERROR(%s)', $start_cell;
    }

    if ( $param->{type} eq 'notContainsErrors' ) {
        $param->{formula} = sprintf 'NOT(ISERROR(%s))', $start_cell;
    }


    # Special handling for 2 color scale.
    if ( $param->{type} eq '2_color_scale' ) {
        $param->{type} = 'colorScale';

        # Color scales don't use any additional formatting.
        $param->{format} = undef;

        # Turn off 3 color parameters.
        $param->{mid_type}  = undef;
        $param->{mid_color} = undef;

        $param->{min_type}  ||= 'min';
        $param->{max_type}  ||= 'max';
        $param->{min_value} ||= 0;
        $param->{max_value} ||= 0;
        $param->{min_color} ||= '#FF7128';
        $param->{max_color} ||= '#FFEF9C';

        $param->{max_color} = $self->_get_palette_color( $param->{max_color} );
        $param->{min_color} = $self->_get_palette_color( $param->{min_color} );
    }


    # Special handling for 3 color scale.
    if ( $param->{type} eq '3_color_scale' ) {
        $param->{type} = 'colorScale';

        # Color scales don't use any additional formatting.
        $param->{format} = undef;

        $param->{min_type}  ||= 'min';
        $param->{mid_type}  ||= 'percentile';
        $param->{max_type}  ||= 'max';
        $param->{min_value} ||= 0;
        $param->{mid_value} = 50 unless defined $param->{mid_value};
        $param->{max_value} ||= 0;
        $param->{min_color} ||= '#F8696B';
        $param->{mid_color} ||= '#FFEB84';
        $param->{max_color} ||= '#63BE7B';

        $param->{max_color} = $self->_get_palette_color( $param->{max_color} );
        $param->{mid_color} = $self->_get_palette_color( $param->{mid_color} );
        $param->{min_color} = $self->_get_palette_color( $param->{min_color} );
    }


    # Special handling for data bar.
    if ( $param->{type} eq 'dataBar' ) {

        # Color scales don't use any additional formatting.
        $param->{format} = undef;

        $param->{min_type}  ||= 'min';
        $param->{max_type}  ||= 'max';
        $param->{min_value} ||= 0;
        $param->{max_value} ||= 0;
        $param->{bar_color} ||= '#638EC6';

        $param->{bar_color} = $self->_get_palette_color( $param->{bar_color} );
    }


    # Store the validation information until we close the worksheet.
    push @{ $self->{_cond_formats}->{$range} }, $param;
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _get_palette_color()
#
# Convert from an Excel internal colour index to a XML style #RRGGBB index
# based on the default or user defined values in the Workbook palette.
#
sub _get_palette_color {

    my $self    = shift;
    my $index   = shift;
    my $palette = $self->{_palette};

    # Handle colours in #XXXXXX RGB format.
    if ( $index =~ m/^#([0-9A-F]{6})$/i ) {
        return "FF" . uc( $1 );
    }

    # Adjust the colour index.
    $index -= 8;

    # Palette is passed in from the Workbook class.
    my @rgb = @{ $palette->[$index] };

    return sprintf "FF%02X%02X%02X", @rgb;
}


###############################################################################
#
# _quote_sheetname()
#
# Sheetnames used in references should be quoted if they contain any spaces,
# special characters or if the look like something that isn't a sheet name.
# TODO. We need to handle more special cases.
#
sub _quote_sheetname {

    my $self      = shift;
    my $sheetname = $_[0];

    if ( $sheetname =~ /^Sheet\d+$/ ) {
        return $sheetname;
    }
    else {
        return qq('$sheetname');
    }
}


###############################################################################
#
# _substitute_cellref()
#
# Substitute an Excel cell reference in A1 notation for  zero based row and
# column values in an argument list.
#
# Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
#
sub _substitute_cellref {

    my $self = shift;
    my $cell = uc( shift );

    # Convert a column range: 'A:A' or 'B:G'.
    # A range such as A:A is equivalent to A1:Rowmax, so add rows as required
    if ( $cell =~ /\$?([A-Z]{1,3}):\$?([A-Z]{1,3})/ ) {
        my ( $row1, $col1 ) = $self->_cell_to_rowcol( $1 . '1' );
        my ( $row2, $col2 ) =
          $self->_cell_to_rowcol( $2 . $self->{_xls_rowmax} );
        return $row1, $col1, $row2, $col2, @_;
    }

    # Convert a cell range: 'A1:B7'
    if ( $cell =~ /\$?([A-Z]{1,3}\$?\d+):\$?([A-Z]{1,3}\$?\d+)/ ) {
        my ( $row1, $col1 ) = $self->_cell_to_rowcol( $1 );
        my ( $row2, $col2 ) = $self->_cell_to_rowcol( $2 );
        return $row1, $col1, $row2, $col2, @_;
    }

    # Convert a cell reference: 'A1' or 'AD2000'
    if ( $cell =~ /\$?([A-Z]{1,3}\$?\d+)/ ) {
        my ( $row1, $col1 ) = $self->_cell_to_rowcol( $1 );
        return $row1, $col1, @_;

    }

    croak( "Unknown cell reference $cell" );
}


###############################################################################
#
# _cell_to_rowcol($cell_ref)
#
# Convert an Excel cell reference in A1 notation to a zero based row and column
# reference; converts C1 to (0, 2).
#
# See also: http://www.perlmonks.org/index.pl?node_id=270352
#
# Returns: ($row, $col, $row_absolute, $col_absolute)
#
#
sub _cell_to_rowcol {

    my $self = shift;

    my $cell = $_[0];
    $cell =~ /(\$?)([A-Z]{1,3})(\$?)(\d+)/;

    my $col_abs = $1 eq "" ? 0 : 1;
    my $col     = $2;
    my $row_abs = $3 eq "" ? 0 : 1;
    my $row     = $4;

    # Convert base26 column string to number
    # All your Base are belong to us.
    my @chars = split //, $col;
    my $expn = 0;
    $col = 0;

    while ( @chars ) {
        my $char = pop( @chars );    # LS char first
        $col += ( ord( $char ) - ord( 'A' ) + 1 ) * ( 26**$expn );
        $expn++;
    }

    # Convert 1-index to zero-index
    $row--;
    $col--;

    # TODO Check row and column range
    return $row, $col, $row_abs, $col_abs;
}


###############################################################################
#
# _sort_pagebreaks()
#
# This is an internal method that is used to filter elements of the array of
# pagebreaks used in the _store_hbreak() and _store_vbreak() methods. It:
#   1. Removes duplicate entries from the list.
#   2. Sorts the list.
#   3. Removes 0 from the list if present.
#
sub _sort_pagebreaks {

    my $self = shift;

    return () unless @_;

    my %hash;
    my @array;

    @hash{@_} = undef;    # Hash slice to remove duplicates
    @array = sort { $a <=> $b } keys %hash;    # Numerical sort
    shift @array if $array[0] == 0;            # Remove zero

    # The Excel 2007 specification says that the maximum number of page breaks
    # is 1026. However, in practice it is actually 1023.
    my $max_num_breaks = 1023;
    splice( @array, $max_num_breaks ) if @array > $max_num_breaks;

    return @array;
}


###############################################################################
#
# _check_dimensions($row, $col, $ignore_row, $ignore_col)
#
# Check that $row and $col are valid and store max and min values for use in
# other methods/elements.
#
# The $ignore_row/$ignore_col flags is used to indicate that we wish to
# perform the dimension check without storing the value.
#
# The ignore flags are use by set_row() and data_validate.
#
sub _check_dimensions {

    my $self       = shift;
    my $row        = $_[0];
    my $col        = $_[1];
    my $ignore_row = $_[2];
    my $ignore_col = $_[3];


    return -2 if not defined $row;
    return -2 if $row >= $self->{_xls_rowmax};

    return -2 if not defined $col;
    return -2 if $col >= $self->{_xls_colmax};

    # In optimization mode we don't change dimensions for rows that are
    # already written.
    if ( !$ignore_row && !$ignore_col && $self->{_optimization} == 1 ) {
        return -2 if $row < $self->{_previous_row};
    }

    if ( !$ignore_row ) {

        if ( not defined $self->{_dim_rowmin} or $row < $self->{_dim_rowmin} ) {
            $self->{_dim_rowmin} = $row;
        }

        if ( not defined $self->{_dim_rowmax} or $row > $self->{_dim_rowmax} ) {
            $self->{_dim_rowmax} = $row;
        }
    }

    if ( !$ignore_col ) {

        if ( not defined $self->{_dim_colmin} or $col < $self->{_dim_colmin} ) {
            $self->{_dim_colmin} = $col;
        }

        if ( not defined $self->{_dim_colmax} or $col > $self->{_dim_colmax} ) {
            $self->{_dim_colmax} = $col;
        }
    }

    return 0;
}


###############################################################################
#
#  _position_object_pixels()
#
# Calculate the vertices that define the position of a graphical object within
# the worksheet in pixels.
#
#         +------------+------------+
#         |     A      |      B     |
#   +-----+------------+------------+
#   |     |(x1,y1)     |            |
#   |  1  |(A1)._______|______      |
#   |     |    |              |     |
#   |     |    |              |     |
#   +-----+----|    BITMAP    |-----+
#   |     |    |              |     |
#   |  2  |    |______________.     |
#   |     |            |        (B2)|
#   |     |            |     (x2,y2)|
#   +---- +------------+------------+
#
# Example of an object that covers some of the area from cell A1 to cell B2.
#
# Based on the width and height of the object we need to calculate 8 vars:
#
#     $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
#
# We also calculate the absolute x and y position of the top left vertex of
# the object. This is required for images.
#
#    $x_abs, $y_abs
#
# The width and height of the cells that the object occupies can be variable
# and have to be taken into account.
#
# The values of $col_start and $row_start are passed in from the calling
# function. The values of $col_end and $row_end are calculated by subtracting
# the width and height of the object from the width and height of the
# underlying cells.
#
sub _position_object_pixels {

    my $self = shift;

    my $col_start;    # Col containing upper left corner of object.
    my $x1;           # Distance to left side of object.

    my $row_start;    # Row containing top left corner of object.
    my $y1;           # Distance to top of object.

    my $col_end;      # Col containing lower right corner of object.
    my $x2;           # Distance to right side of object.

    my $row_end;      # Row containing bottom right corner of object.
    my $y2;           # Distance to bottom of object.

    my $width;        # Width of object frame.
    my $height;       # Height of object frame.

    my $x_abs = 0;    # Absolute distance to left side of object.
    my $y_abs = 0;    # Absolute distance to top  side of object.

    my $is_drawing = 0;

    ( $col_start, $row_start, $x1, $y1, $width, $height, $is_drawing) = @_;

    # Calculate the absolute x offset of the top-left vertex.
    if ( $self->{_col_size_changed} ) {
        for my $col_id ( 1 .. $col_start ) {
            $x_abs += $self->_size_col( $col_id );
        }
    }
    else {
        # Optimisation for when the column widths haven't changed.
        $x_abs += 64 * $col_start;
    }

    $x_abs += $x1;

    # Calculate the absolute y offset of the top-left vertex.
    # Store the column change to allow optimisations.
    if ( $self->{_row_size_changed} ) {
        for my $row_id ( 1 .. $row_start ) {
            $y_abs += $self->_size_row( $row_id );
        }
    }
    else {
        # Optimisation for when the row heights haven't changed.
        $y_abs += 20 * $row_start;
    }

    $y_abs += $y1;


    # Adjust start column for offsets that are greater than the col width.
    while ( $x1 >= $self->_size_col( $col_start ) ) {
        $x1 -= $self->_size_col( $col_start );
        $col_start++;
    }

    # Adjust start row for offsets that are greater than the row height.
    while ( $y1 >= $self->_size_row( $row_start ) ) {
        $y1 -= $self->_size_row( $row_start );
        $row_start++;
    }


    # Initialise end cell to the same as the start cell.
    $col_end = $col_start;
    $row_end = $row_start;

    $width  = $width + $x1;
    $height = $height + $y1;


    # Subtract the underlying cell widths to find the end cell of the object.
    while ( $width >= $self->_size_col( $col_end ) ) {
        $width -= $self->_size_col( $col_end );
        $col_end++;
    }


    # Subtract the underlying cell heights to find the end cell of the object.
    while ( $height >= $self->_size_row( $row_end ) ) {
        $height -= $self->_size_row( $row_end );
        $row_end++;
    }

    # The following is only required for positioning drawing/chart objects
    # and not comments. It is probably the result of a bug.
    if ( $is_drawing ) {
        $col_end-- if $width == 0;
        $row_end-- if $height == 0;
    }

    # The end vertices are whatever is left from the width and height.
    $x2 = $width;
    $y2 = $height;

    return (
        $col_start, $row_start, $x1, $y1,
        $col_end,   $row_end,   $x2, $y2,
        $x_abs,     $y_abs

    );
}


###############################################################################
#
#  _position_object_emus()
#
# Calculate the vertices that define the position of a graphical object within
# the worksheet in EMUs.
#
# The vertices are expressed as English Metric Units (EMUs). There are 12,700
# EMUs per point. Therefore, 12,700 * 3 /4 = 9,525 EMUs per pixel.
#
sub _position_object_emus {

    my $self       = shift;
    my $is_drawing = 1;

    my (
        $col_start, $row_start, $x1, $y1,
        $col_end,   $row_end,   $x2, $y2,
        $x_abs,     $y_abs

    ) = $self->_position_object_pixels( @_, $is_drawing );

    # Convert the pixel values to EMUs. See above.
    $x1    *= 9_525;
    $y1    *= 9_525;
    $x2    *= 9_525;
    $y2    *= 9_525;
    $x_abs *= 9_525;
    $y_abs *= 9_525;

    return (
        $col_start, $row_start, $x1, $y1,
        $col_end,   $row_end,   $x2, $y2,
        $x_abs,     $y_abs

    );
}


###############################################################################
#
# _size_col($col)
#
# Convert the width of a cell from user's units to pixels. Excel rounds the
# column width to the nearest pixel. If the width hasn't been set by the user
# we use the default value. If the column is hidden it has a value of zero.
#
sub _size_col {

    my $self = shift;
    my $col  = shift;

    my $max_digit_width = 7;    # For Calabri 11.
    my $padding         = 5;
    my $pixels;

    # Look up the cell value to see if it has been changed.
    if ( exists $self->{_col_sizes}->{$col} ) {
        my $width = $self->{_col_sizes}->{$col};

        # Convert to pixels.
        if ( $width == 0) {
            $pixels = 0;
        }
        elsif ( $width < 1 ) {
            $pixels = int( $width * 12 + 0.5 );
        }
        else {
            $pixels = int( $width * $max_digit_width + 0.5 ) + $padding;
        }
    }
    else {
        $pixels = 64;
    }

    return $pixels;
}


###############################################################################
#
# _size_row($row)
#
# Convert the height of a cell from user's units to pixels. If the height
# hasn't been set by the user we use the default value. If the row is hidden
# it has a value of zero.
#
sub _size_row {

    my $self = shift;
    my $row  = shift;
    my $pixels;

    # Look up the cell value to see if it has been changed
    if ( exists $self->{_row_sizes}->{$row} ) {
        my $height = $self->{_row_sizes}->{$row};

        if ( $height == 0 ) {
            $pixels = 0;
        }
        else {
            $pixels = int( 4 / 3 * $height );
        }
    }
    else {
        $pixels = 20;
    }

    return $pixels;
}


###############################################################################
#
# _options_changed()
#
# Check to see if any of the worksheet options have changed.
#
sub _options_changed {

    my $self = shift;

    my $options_changed = 0;
    my $print_changed   = 0;
    my $setup_changed   = 0;


    if (   $self->{_orientation} == 0
        or $self->{_hcenter} == 1
        or $self->{_vcenter} == 1
        or $self->{_header} ne ''
        or $self->{_footer} ne ''
        or $self->{_margin_header} != 0.50
        or $self->{_margin_footer} != 0.50
        or $self->{_margin_left} != 0.75
        or $self->{_margin_right} != 0.75
        or $self->{_margin_top} != 1.00
        or $self->{_margin_bottom} != 1.00 )
    {
        $setup_changed = 1;
    }


    # Special case for 1x1 page fit.
    if ( $self->{_fit_width} == 1 and $self->{_fit_height} == 1 ) {
        $options_changed     = 1;
        $self->{_fit_width}  = 0;
        $self->{_fit_height} = 0;
    }


    if (   $self->{_fit_width} > 1
        or $self->{_fit_height} > 1
        or $self->{_page_order} == 1
        or $self->{_black_white} == 1
        or $self->{_draft_quality} == 1
        or $self->{_print_comments} == 1
        or $self->{_paper_size} != 0
        or $self->{_print_scale} != 100
        or $self->{_print_gridlines} == 1
        or $self->{_print_headers} == 1
        or @{ $self->{_hbreaks} } > 0
        or @{ $self->{_vbreaks} } > 0 )
    {
        $print_changed = 1;
    }


    if (   $print_changed
        or $setup_changed )
    {
        $options_changed = 1;
    }


    $options_changed = 1 if $self->{_screen_gridlines} == 0;
    $options_changed = 1 if $self->{_filter_on};

    return ( $options_changed, $print_changed, $setup_changed );
}


###############################################################################
#
# _get_shared_string_index()
#
# Add a string to the shared string table, if it isn't already there, and
# return the string index.
#
sub _get_shared_string_index {

    my $self = shift;
    my $str  = shift;

    # Add the string to the shared string table.
    if ( not exists ${ $self->{_str_table} }->{$str} ) {
        ${ $self->{_str_table} }->{$str} = ${ $self->{_str_unique} }++;
    }

    ${ $self->{_str_total} }++;
    my $index = ${ $self->{_str_table} }->{$str};

    return $index;
}


###############################################################################
#
# insert_chart( $row, $col, $chart, $x, $y, $scale_x, $scale_y )
#
# Insert a chart into a worksheet. The $chart argument should be a Chart
# object or else it is assumed to be a filename of an external binary file.
# The latter is for backwards compatibility.
#
sub insert_chart {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column.
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    my $row      = $_[0];
    my $col      = $_[1];
    my $chart    = $_[2];
    my $x_offset = $_[3] || 0;
    my $y_offset = $_[4] || 0;
    my $scale_x  = $_[5] || 1;
    my $scale_y  = $_[6] || 1;

    croak "Insufficient arguments in insert_chart()" unless @_ >= 3;

    if ( ref $chart ) {

        # Check for a Chart object.
        croak "Not a Chart object in insert_chart()"
          unless $chart->isa( 'Excel::Writer::XLSX::Chart' );

        # Check that the chart is an embedded style chart.
        croak "Not a embedded style Chart object in insert_chart()"
          unless $chart->{_embedded};

    }

    push @{ $self->{_charts} },
      [ $row, $col, $chart, $x_offset, $y_offset, $scale_x, $scale_y ];
}


###############################################################################
#
# _prepare_chart()
#
# Set up chart/drawings.
#
sub _prepare_chart {

    my $self         = shift;
    my $index        = shift;
    my $chart_id     = shift;
    my $drawing_id   = shift;
    my $drawing_type = 1;

    my ( $row, $col, $chart, $x_offset, $y_offset, $scale_x, $scale_y ) =
      @{ $self->{_charts}->[$index] };

    my $width  = int( 0.5 + ( 480 * $scale_x ) );
    my $height = int( 0.5 + ( 288 * $scale_y ) );

    my @dimensions =
      $self->_position_object_emus( $col, $row, $x_offset, $y_offset, $width,
        $height );

    # Create a Drawing object to use with worksheet unless one already exists.
    if ( !$self->{_drawing} ) {

        my $drawing = Excel::Writer::XLSX::Drawing->new();
        $drawing->_add_drawing_object( $drawing_type, @dimensions );
        $drawing->{_embedded} = 1;

        $self->{_drawing} = $drawing;

        push @{ $self->{_external_drawing_links} },
          [ '/drawing', '../drawings/drawing' . $drawing_id . '.xml' ];
    }
    else {
        my $drawing = $self->{_drawing};
        $drawing->_add_drawing_object( $drawing_type, @dimensions );

    }

    push @{ $self->{_drawing_links} },
      [ '/chart', '../charts/chart' . $chart_id . '.xml' ];
}



###############################################################################
#
# _get_range_data
#
# Returns a range of data from the worksheet _table to be used in chart
# cached data. Strings are returned as SST ids and decoded in the workbook.
# Return undefs for data that doesn't exist since Excel can chart series
# with data missing.
#
sub _get_range_data {

    my $self = shift;

    return () if $self->{_optimization};

    my @data;
    my ( $row_start, $col_start, $row_end, $col_end ) = @_;

    # TODO. Check for worksheet limits.

    # Iterate through the table data.
    for my $row_num ( $row_start .. $row_end ) {

        # Store undef if row doesn't exist.
        if ( !$self->{_table}->[$row_num] ) {
            push @data, undef;
            next;
        }

        for my $col_num ( $col_start .. $col_end ) {

            if ( my $cell = $self->{_table}->[$row_num]->[$col_num] ) {

                my $type  = $cell->[0];
                my $token = $cell->[1];


                if ( $type eq 'n' ) {

                    # Store a number.
                    push @data, $token;
                }
                elsif ( $type eq 's' ) {

                    # Store a string.
                    if ( $self->{_optimization} == 0 ) {
                        push @data, { 'sst_id' => $token};
                    }
                    else {
                        push @data, $token;
                    }
                }
                elsif ( $type eq 'f' ) {

                    # Store a formula.
                    push @data, $cell->[3] || 0;
                }
                elsif ( $type eq 'a' ) {

                    # Store an array formula.
                    push @data, $cell->[4] || 0;
                }
                elsif ( $type eq 'l' ) {

                    # Store the string part a hyperlink.
                    push @data, { 'sst_id' => $token};
                }
                elsif ( $type eq 'b' ) {

                    # Store a empty cell.
                    push @data, '';
                }
            }
            else {

                # Store undef if col doesn't exist.
                push @data, undef;
            }
        }
    }

    return @data;
}


###############################################################################
#
# insert_image( $row, $col, $filename, $x, $y, $scale_x, $scale_y )
#
# Insert an image into the worksheet.
#
sub insert_image {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column.
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    my $row      = $_[0];
    my $col      = $_[1];
    my $image    = $_[2];
    my $x_offset = $_[3] || 0;
    my $y_offset = $_[4] || 0;
    my $scale_x  = $_[5] || 1;
    my $scale_y  = $_[6] || 1;

    croak "Insufficient arguments in insert_image()" unless @_ >= 3;
    croak "Couldn't locate $image: $!" unless -e $image;

    push @{ $self->{_images} },
      [ $row, $col, $image, $x_offset, $y_offset, $scale_x, $scale_y ];
}


###############################################################################
#
# _prepare_image()
#
# Set up image/drawings.
#
sub _prepare_image {

    my $self         = shift;
    my $index        = shift;
    my $image_id     = shift;
    my $drawing_id   = shift;
    my $width        = shift;
    my $height       = shift;
    my $name         = shift;
    my $image_type   = shift;
    my $drawing_type = 2;
    my $drawing;

    my ( $row, $col, $image, $x_offset, $y_offset, $scale_x, $scale_y ) =
      @{ $self->{_images}->[$index] };

    $width  *= $scale_x;
    $height *= $scale_y;

    my @dimensions =
      $self->_position_object_emus( $col, $row, $x_offset, $y_offset, $width,
        $height );

    # Convert from pixels to emus.
    $width  = int( 0.5 + ( $width * 9_525 ) );
    $height = int( 0.5 + ( $height * 9_525 ) );

    # Create a Drawing object to use with worksheet unless one already exists.
    if ( !$self->{_drawing} ) {

        $drawing = Excel::Writer::XLSX::Drawing->new();
        $drawing->{_embedded} = 1;

        $self->{_drawing} = $drawing;

        push @{ $self->{_external_drawing_links} },
          [ '/drawing', '../drawings/drawing' . $drawing_id . '.xml' ];
    }
    else {
        $drawing = $self->{_drawing};
    }

    $drawing->_add_drawing_object( $drawing_type, @dimensions, $width, $height,
        $name );


    push @{ $self->{_drawing_links} },
      [ '/image', '../media/image' .  $image_id . '.' . $image_type ];
}


###############################################################################
#
# _prepare_comments()
#
# Turn the HoH that stores the comments into an array for easier handling
# and set the external links.
#
sub _prepare_comments {

    my $self         = shift;
    my $vml_data_id  = shift;
    my $vml_shape_id = shift;
    my $comment_id   = shift;
    my @comments;

    # We sort the comments by row and column but that isn't strictly required.
    my @rows = sort { $a <=> $b } keys %{ $self->{_comments} };

    for my $row ( @rows ) {
        my @cols = sort { $a <=> $b } keys %{ $self->{_comments}->{$row} };

        for my $col ( @cols ) {

            # Set comment visibility if required and not already user defined.
            if ( $self->{_comments_visible} ) {
                if ( !defined $self->{_comments}->{$row}->{$col}->[4] ) {
                    $self->{_comments}->{$row}->{$col}->[4] = 1;
                }
            }

            # Set comment author if not already user defined.
            if ( !defined $self->{_comments}->{$row}->{$col}->[3] ) {
                $self->{_comments}->{$row}->{$col}->[3] =
                  $self->{_comments_author};
            }

            push @comments, $self->{_comments}->{$row}->{$col};
        }
    }

    $self->{_comments_array} = \@comments;

    push @{ $self->{_external_comment_links} },
      [ '/vmlDrawing', '../drawings/vmlDrawing' . $comment_id . '.vml' ],
      [ '/comments',   '../comments' . $comment_id . '.xml' ];

    my $count         = scalar @comments;
    my $start_data_id = $vml_data_id;

    # The VML o:idmap data id contains a comma separated range when there is
    # more than one 1024 block of comments, like this: data="1,2".
    for my $i ( 1 .. int( $count / 1024 ) ) {
        $vml_data_id = "$vml_data_id," . ( $start_data_id + $i );
    }

    $self->{_vml_data_id}  = $vml_data_id;
    $self->{_vml_shape_id} = $vml_shape_id;

    return $count;
}


###############################################################################
#
# _comment_params()
#
# This method handles the additional optional parameters to write_comment() as
# well as calculating the comment object position and vertices.
#
sub _comment_params {

    my $self = shift;

    my $row    = shift;
    my $col    = shift;
    my $string = shift;

    my $default_width  = 128;
    my $default_height = 74;

    my %params = (
        author          => undef,
        color           => 81,
        start_cell      => undef,
        start_col       => undef,
        start_row       => undef,
        visible         => undef,
        width           => $default_width,
        height          => $default_height,
        x_offset        => undef,
        x_scale         => 1,
        y_offset        => undef,
        y_scale         => 1,
    );


    # Overwrite the defaults with any user supplied values. Incorrect or
    # misspelled parameters are silently ignored.
    %params = ( %params, @_ );


    # Ensure that a width and height have been set.
    $params{width}  = $default_width  if not $params{width};
    $params{height} = $default_height if not $params{height};


    # Limit the string to the max number of chars.
    my $max_len = 32767;

    if ( length( $string ) > $max_len ) {
        $string = substr( $string, 0, $max_len );
    }


    # Set the comment background colour.
    my $color    = $params{color};
    my $color_id = &Excel::Writer::XLSX::Format::_get_color( $color );

    if ( $color_id == 0 ) {
        $params{color} = '#ffffe1';
    }
    else {
        my $palette = $self->{_palette};

        # Get the RGB color from the palette.
        my @rgb = @{ $palette->[ $color_id - 8 ] };
        my $rgb_color = sprintf "%02x%02x%02x", @rgb;

        # Minor modification to allow comparison testing. Change RGB colors
        # from long format, ffcc00 to short format fc0 used by VML.
        $rgb_color =~ s/^([0-9a-f])\1([0-9a-f])\2([0-9a-f])\3$/$1$2$3/;

        $params{color} = sprintf "#%s [%d]\n", $rgb_color, $color_id;
    }


    # Convert a cell reference to a row and column.
    if ( defined $params{start_cell} ) {
        my ( $row, $col ) = $self->_substitute_cellref( $params{start_cell} );
        $params{start_row} = $row;
        $params{start_col} = $col;
    }


    # Set the default start cell and offsets for the comment. These are
    # generally fixed in relation to the parent cell. However there are
    # some edge cases for cells at the, er, edges.
    #
    my $row_max = $self->{_xls_rowmax};
    my $col_max = $self->{_xls_colmax};

    if ( not defined $params{start_row} ) {

        if    ( $row == 0 )            { $params{start_row} = 0 }
        elsif ( $row == $row_max - 3 ) { $params{start_row} = $row_max - 7 }
        elsif ( $row == $row_max - 2 ) { $params{start_row} = $row_max - 6 }
        elsif ( $row == $row_max - 1 ) { $params{start_row} = $row_max - 5 }
        else                           { $params{start_row} = $row - 1 }
    }

    if ( not defined $params{y_offset} ) {

        if    ( $row == 0 )            { $params{y_offset} = 2 }
        elsif ( $row == $row_max - 3 ) { $params{y_offset} = 16 }
        elsif ( $row == $row_max - 2 ) { $params{y_offset} = 16 }
        elsif ( $row == $row_max - 1 ) { $params{y_offset} = 14 }
        else                           { $params{y_offset} = 10 }
    }

    if ( not defined $params{start_col} ) {

        if    ( $col == $col_max - 3 ) { $params{start_col} = $col_max - 6 }
        elsif ( $col == $col_max - 2 ) { $params{start_col} = $col_max - 5 }
        elsif ( $col == $col_max - 1 ) { $params{start_col} = $col_max - 4 }
        else                           { $params{start_col} = $col + 1 }
    }

    if ( not defined $params{x_offset} ) {

        if    ( $col == $col_max - 3 ) { $params{x_offset} = 49 }
        elsif ( $col == $col_max - 2 ) { $params{x_offset} = 49 }
        elsif ( $col == $col_max - 1 ) { $params{x_offset} = 49 }
        else                           { $params{x_offset} = 15 }
    }


    # Scale the size of the comment box if required.
    if ( $params{x_scale} ) {
        $params{width} = $params{width} * $params{x_scale};
    }

    if ( $params{y_scale} ) {
        $params{height} = $params{height} * $params{y_scale};
    }

    # Round the dimensions to the nearest pixel.
    $params{width}  = int( 0.5 + $params{width} );
    $params{height} = int( 0.5 + $params{height} );

    # Calculate the positions of comment object.
    my @vertices = $self->_position_object_pixels(
        $params{start_col}, $params{start_row}, $params{x_offset},
        $params{y_offset},  $params{width},     $params{height}
      );

    # Add the width and height for VML.
    push @vertices, ( $params{width}, $params{height} );

    return (
        $row,
        $col,
        $string,

        $params{author},
        $params{visible},
        $params{color},

        [@vertices]
    );
}


###############################################################################
#
# Deprecated methods for backwards compatibility.
#
###############################################################################


# This method was mainly only required for Excel 5.
sub write_url_range { }

# Deprecated UTF-16 method required for the Excel 5 format.
sub write_utf16be_string {

    my $self = shift;

    # Convert A1 notation if present.
    @_ = $self->_substitute_cellref( @_ ) if $_[0] =~ /^\D/;

    # Check the number of args.
    return -1 if @_ < 3;

    # Convert UTF16 string to UTF8.
    require Encode;
    my $utf8_string = Encode::decode( 'UTF-16BE', $_[2] );

    return $self->write_string( $_[0], $_[1], $utf8_string, $_[3] );
}

# Deprecated UTF-16 method required for the Excel 5 format.
sub write_utf16le_string {

    my $self = shift;

    # Convert A1 notation if present.
    @_ = $self->_substitute_cellref( @_ ) if $_[0] =~ /^\D/;

    # Check the number of args.
    return -1 if @_ < 3;

    # Convert UTF16 string to UTF8.
    require Encode;
    my $utf8_string = Encode::decode( 'UTF-16LE', $_[2] );

    return $self->write_string( $_[0], $_[1], $utf8_string, $_[3] );
}

# No longer required. Was used to avoid slow formula parsing.
sub store_formula {

    my $self   = shift;
    my $string = shift;

    my @tokens = split /(\$?[A-I]?[A-Z]\$?\d+)/, $string;

    return \@tokens;
}

# No longer required. Was used to avoid slow formula parsing.
sub repeat_formula {

    my $self = shift;

    # Convert A1 notation if present.
    @_ = $self->_substitute_cellref( @_ ) if $_[0] =~ /^\D/;

    if ( @_ < 2 ) { return -1 }    # Check the number of args

    my $row         = shift;       # Zero indexed row
    my $col         = shift;       # Zero indexed column
    my $formula_ref = shift;       # Array ref with formula tokens
    my $format      = shift;       # XF format
    my @pairs       = @_;          # Pattern/replacement pairs


    # Enforce an even number of arguments in the pattern/replacement list.
    croak "Odd number of elements in pattern/replacement list" if @pairs % 2;

    # Check that $formula is an array ref.
    croak "Not a valid formula" if ref $formula_ref ne 'ARRAY';

    my @tokens = @$formula_ref;

    # Allow the user to specify the result of the formula by appending a
    # result => $value pair to the end of the arguments.
    my $value = undef;
    if ( @pairs && $pairs[-2] eq 'result' ) {
        $value = pop @pairs;
        pop @pairs;
    }

    # Make the substitutions.
    while ( @pairs ) {
        my $pattern = shift @pairs;
        my $replace = shift @pairs;

        foreach my $token ( @tokens ) {
            last if $token =~ s/$pattern/$replace/;
        }
    }

    my $formula = join '', @tokens;

    return $self->write_formula( $row, $col, $formula, $format, $value );
}


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_worksheet()
#
# Write the <worksheet> element. This is the root element of Worksheet.
#
sub _write_worksheet {

    my $self                   = shift;
    my $schema                 = 'http://schemas.openxmlformats.org/';
    my $xmlns                  = $schema . 'spreadsheetml/2006/main';
    my $xmlns_r                = $schema . 'officeDocument/2006/relationships';
    my $xmlns_mc               = $schema . 'markup-compatibility/2006';
    my $xmlns_mv               = 'urn:schemas-microsoft-com:mac:vml';
    my $mc_ignorable           = 'mv';
    my $mc_preserve_attributes = 'mv:*';

    my @attributes = (
        'xmlns'   => $xmlns,
        'xmlns:r' => $xmlns_r,
    );

    $self->{_writer}->startTag( 'worksheet', @attributes );
}


###############################################################################
#
# _write_sheet_pr()
#
# Write the <sheetPr> element for Sheet level properties.
#
sub _write_sheet_pr {

    my $self       = shift;
    my @attributes = ();

    if (   !$self->{_fit_page}
        && !$self->{_filter_on}
        && !$self->{_tab_color}
        && !$self->{_outline_changed} )
    {
        return;
    }

    push @attributes, ( 'filterMode' => 1 ) if $self->{_filter_on};

    if (   $self->{_fit_page}
        || $self->{_tab_color}
        || $self->{_outline_changed} )
    {
        $self->{_writer}->startTag( 'sheetPr', @attributes );
        $self->_write_tab_color();
        $self->_write_outline_pr();
        $self->_write_page_set_up_pr();
        $self->{_writer}->endTag( 'sheetPr' );
    }
    else {
        $self->{_writer}->emptyTag( 'sheetPr', @attributes );
    }
}


##############################################################################
#
# _write_page_set_up_pr()
#
# Write the <pageSetUpPr> element.
#
sub _write_page_set_up_pr {

    my $self = shift;

    return unless $self->{_fit_page};

    my @attributes = ( 'fitToPage' => 1 );

    $self->{_writer}->emptyTag( 'pageSetUpPr', @attributes );
}


###############################################################################
#
# _write_dimension()
#
# Write the <dimension> element. This specifies the range of cells in the
# worksheet. As a special case, empty spreadsheets use 'A1' as a range.
#
sub _write_dimension {

    my $self = shift;
    my $ref;

    if ( !defined $self->{_dim_rowmin} && !defined $self->{_dim_colmin} ) {

        # If the min dims are undefined then no dimensions have been set
        # and we use the default 'A1'.
        $ref = 'A1';
    }
    elsif ( !defined $self->{_dim_rowmin} && defined $self->{_dim_colmin} ) {

        # If the row dims aren't set but the column dims are then they
        # have been changed via set_column().

        if ( $self->{_dim_colmin} == $self->{_dim_colmax} ) {

            # The dimensions are a single cell and not a range.
            $ref = xl_rowcol_to_cell( 0, $self->{_dim_colmin} );
        }
        else {

            # The dimensions are a cell range.
            my $cell_1 = xl_rowcol_to_cell( 0, $self->{_dim_colmin} );
            my $cell_2 = xl_rowcol_to_cell( 0, $self->{_dim_colmax} );

            $ref = $cell_1 . ':' . $cell_2;
        }

    }
    elsif ($self->{_dim_rowmin} == $self->{_dim_rowmax}
        && $self->{_dim_colmin} == $self->{_dim_colmax} )
    {

        # The dimensions are a single cell and not a range.
        $ref = xl_rowcol_to_cell( $self->{_dim_rowmin}, $self->{_dim_colmin} );
    }
    else {

        # The dimensions are a cell range.
        my $cell_1 =
          xl_rowcol_to_cell( $self->{_dim_rowmin}, $self->{_dim_colmin} );
        my $cell_2 =
          xl_rowcol_to_cell( $self->{_dim_rowmax}, $self->{_dim_colmax} );

        $ref = $cell_1 . ':' . $cell_2;
    }


    my @attributes = ( 'ref' => $ref );

    $self->{_writer}->emptyTag( 'dimension', @attributes );
}


###############################################################################
#
# _write_sheet_views()
#
# Write the <sheetViews> element.
#
sub _write_sheet_views {

    my $self = shift;

    my @attributes = ();

    $self->{_writer}->startTag( 'sheetViews', @attributes );
    $self->_write_sheet_view();
    $self->{_writer}->endTag( 'sheetViews' );
}


###############################################################################
#
# _write_sheet_view()
#
# Write the <sheetView> element.
#
# Sample structure:
#     <sheetView
#         showGridLines="0"
#         showRowColHeaders="0"
#         showZeros="0"
#         rightToLeft="1"
#         tabSelected="1"
#         showRuler="0"
#         showOutlineSymbols="0"
#         view="pageLayout"
#         zoomScale="121"
#         zoomScaleNormal="121"
#         workbookViewId="0"
#      />
#
sub _write_sheet_view {

    my $self             = shift;
    my $gridlines        = $self->{_screen_gridlines};
    my $show_zeros       = $self->{_show_zeros};
    my $right_to_left    = $self->{_right_to_left};
    my $tab_selected     = $self->{_selected};
    my $view             = $self->{_page_view};
    my $zoom             = $self->{_zoom};
    my $workbook_view_id = 0;
    my @attributes       = ();

    # Hide screen gridlines if required
    if ( !$gridlines ) {
        push @attributes, ( 'showGridLines' => 0 );
    }

    # Hide zeroes in cells.
    if ( !$show_zeros ) {
        push @attributes, ( 'showZeros' => 0 );
    }

    # Display worksheet right to left for Hebrew, Arabic and others.
    if ( $right_to_left ) {
        push @attributes, ( 'rightToLeft' => 1 );
    }

    # Show that the sheet tab is selected.
    if ( $tab_selected ) {
        push @attributes, ( 'tabSelected' => 1 );
    }


    # Turn outlines off. Also required in the outlinePr element.
    if ( !$self->{_outline_on} ) {
        push @attributes, ( "showOutlineSymbols" => 0 );
    }

    # Set the page view/layout mode if required.
    # TODO. Add pageBreakPreview mode when requested.
    if ( $view ) {
        push @attributes, ( 'view' => 'pageLayout' );
    }

    # Set the zoom level.
    if ( $zoom != 100 ) {
        push @attributes, ( 'zoomScale' => $zoom ) unless $view;
        push @attributes, ( 'zoomScaleNormal' => $zoom )
          if $self->{_zoom_scale_normal};
    }

    push @attributes, ( 'workbookViewId' => $workbook_view_id );

    if ( @{ $self->{_panes} } || @{ $self->{_selections} } ) {
        $self->{_writer}->startTag( 'sheetView', @attributes );
        $self->_write_panes();
        $self->_write_selections();
        $self->{_writer}->endTag( 'sheetView' );
    }
    else {
        $self->{_writer}->emptyTag( 'sheetView', @attributes );
    }
}


###############################################################################
#
# _write_selections()
#
# Write the <selection> elements.
#
sub _write_selections {

    my $self = shift;

    for my $selection ( @{ $self->{_selections} } ) {
        $self->_write_selection( @$selection );
    }
}


###############################################################################
#
# _write_selection()
#
# Write the <selection> element.
#
sub _write_selection {

    my $self        = shift;
    my $pane        = shift;
    my $active_cell = shift;
    my $sqref       = shift;
    my @attributes  = ();

    push @attributes, ( 'pane'       => $pane )        if $pane;
    push @attributes, ( 'activeCell' => $active_cell ) if $active_cell;
    push @attributes, ( 'sqref'      => $sqref )       if $sqref;

    $self->{_writer}->emptyTag( 'selection', @attributes );
}


###############################################################################
#
# _write_sheet_format_pr()
#
# Write the <sheetFormatPr> element.
#
sub _write_sheet_format_pr {

    my $self               = shift;
    my $base_col_width     = 10;
    my $default_row_height = 15;
    my $row_level      = $self->{_outline_row_level};
    my $col_level      = $self->{_outline_col_level};

    my @attributes = ( 'defaultRowHeight' => $default_row_height );
    push @attributes, ( 'outlineLevelRow' => $row_level ) if $row_level;
    push @attributes, ( 'outlineLevelCol' => $col_level ) if $col_level;

    $self->{_writer}->emptyTag( 'sheetFormatPr', @attributes );
}


##############################################################################
#
# _write_cols()
#
# Write the <cols> element and <col> sub elements.
#
sub _write_cols {

    my $self = shift;

    # Exit unless some column have been formatted.
    return unless @{ $self->{_colinfo} };

    $self->{_writer}->startTag( 'cols' );

    for my $col_info ( @{ $self->{_colinfo} } ) {
        $self->_write_col_info( @$col_info );
    }

    $self->{_writer}->endTag( 'cols' );
}


##############################################################################
#
# _write_col_info()
#
# Write the <col> element.
#
sub _write_col_info {

    my $self         = shift;
    my $min          = $_[0] || 0;    # First formatted column.
    my $max          = $_[1] || 0;    # Last formatted column.
    my $width        = $_[2];         # Col width in user units.
    my $format       = $_[3];         # Format index.
    my $hidden       = $_[4] || 0;    # Hidden flag.
    my $level        = $_[5] || 0;    # Outline level.
    my $collapsed    = $_[6] || 0;    # Outline level.
    my $custom_width = 1;
    my $xf_index     = 0;

    # Get the format index.
    if ( ref( $format ) ) {
        $xf_index =  $format->get_xf_index();
    }

    # Set the Excel default col width.
    if ( !defined $width ) {
        if ( !$hidden ) {
            $width        = 8.43;
            $custom_width = 0;
        }
        else {
            $width = 0;
        }
    }
    else {

        # Width is defined but same as default.
        if ( $width == 8.43 ) {
            $custom_width = 0;
        }
    }


    # Convert column width from user units to character width.
    my $max_digit_width = 7;    # For Calabri 11.
    my $padding         = 5;
    if ( $width > 0 ) {
        $width = int(
            ( $width * $max_digit_width + $padding ) / $max_digit_width * 256 )
          / 256;
    }

    my @attributes = (
        'min'   => $min + 1,
        'max'   => $max + 1,
        'width' => $width,
    );

    push @attributes, ( 'style'        => $xf_index ) if $xf_index;
    push @attributes, ( 'hidden'       => 1 )         if $hidden;
    push @attributes, ( 'customWidth'  => 1 )         if $custom_width;
    push @attributes, ( 'outlineLevel' => $level )    if $level;
    push @attributes, ( 'collapsed'    => 1 )         if $collapsed;


    $self->{_writer}->emptyTag( 'col', @attributes );
}


###############################################################################
#
# _write_sheet_data()
#
# Write the <sheetData> element.
#
sub _write_sheet_data {

    my $self = shift;

    if ( not defined $self->{_dim_rowmin} ) {

        # If the dimensions aren't defined then there is no data to write.
        $self->{_writer}->emptyTag( 'sheetData' );
    }
    else {
        $self->{_writer}->startTag( 'sheetData' );
        $self->_write_rows();
        $self->{_writer}->endTag( 'sheetData' );

    }

}


###############################################################################
#
# _write_optimized_sheet_data()
#
# Write the <sheetData> element when the memory optimisation is on. In which
# case we read the data stored in the temp file and rewrite it to the XML
# sheet file.
#
sub _write_optimized_sheet_data {

    my $self = shift;

    if ( not defined $self->{_dim_rowmin} ) {

        # If the dimensions aren't defined then there is no data to write.
        $self->{_writer}->emptyTag( 'sheetData' );
    }
    else {
        $self->{_writer}->startTag( 'sheetData' );

        my $xlsx_fh = $self->{_writer}->getOutput();
        my $cell_fh = $self->{_cell_data_fh};

        my $buffer;
        # Rewind the temp file.
        seek $cell_fh, 0, 0;

        while ( read( $cell_fh, $buffer, 4_096 ) ) {
            local $\ = undef;    # Protect print from -l on commandline.
            print $xlsx_fh $buffer;
        }

        $self->{_writer}->endTag( 'sheetData' );
    }
}


###############################################################################
#
# _write_rows()
#
# Write out the worksheet data as a series of rows and cells.
#
sub _write_rows {

    my $self = shift;

    $self->_calculate_spans();

    for my $row_num ( $self->{_dim_rowmin} .. $self->{_dim_rowmax} ) {

        # Skip row if it doesn't contain row formatting, cell data or a comment.
        if (   !$self->{_set_rows}->{$row_num}
            && !$self->{_table}->[$row_num]
            && !$self->{_comments}->{$row_num} )
        {
            next;
        }

        my $span_index = int( $row_num / 16 );
        my $span       = $self->{_row_spans}->[$span_index];

        # Write the cells if the row contains data.
        if ( my $row_ref = $self->{_table}->[$row_num] ) {

            if ( !$self->{_set_rows}->{$row_num} ) {
                $self->_write_row( $row_num, $span );
            }
            else {
                $self->_write_row( $row_num, $span,
                    @{ $self->{_set_rows}->{$row_num} } );
            }


            for my $col_num ( $self->{_dim_colmin} .. $self->{_dim_colmax} ) {
                if ( my $col_ref = $self->{_table}->[$row_num]->[$col_num] ) {
                    $self->_write_cell( $row_num, $col_num, $col_ref );
                }
            }

            $self->{_writer}->endTag( 'row' );
        }
        elsif ( $self->{_comments}->{$row_num} ) {

            $self->_write_empty_row( $row_num, $span,
                @{ $self->{_set_rows}->{$row_num} } );
        }
        else {

            # Row attributes only.
            $self->_write_empty_row( $row_num, undef,
                @{ $self->{_set_rows}->{$row_num} } );
        }
    }
}


###############################################################################
#
# _write_single_row()
#
# Write out the worksheet data as a single row with cells. This method is
# used when memory optimisation is on. A single row is written and the data
# table is reset. That way only one row of data is kept in memory at any one
# time. We don't write span data in the optimised case since it is optional.
#
sub _write_single_row {

    my $self        = shift;
    my $current_row = shift || 0;
    my $row_num     = $self->{_previous_row};

    # Set the new previous row as the current row.
    $self->{_previous_row} = $current_row;

    # Skip row if it doesn't contain row formatting, cell data or a comment.
    if (   !$self->{_set_rows}->{$row_num}
        && !$self->{_table}->[$row_num]
        && !$self->{_comments}->{$row_num} )
    {
        return;
    }

    # Write the cells if the row contains data.
    if ( my $row_ref = $self->{_table}->[$row_num] ) {

        if ( !$self->{_set_rows}->{$row_num} ) {
            $self->_write_row( $row_num );
        }
        else {
            $self->_write_row( $row_num, undef,
                @{ $self->{_set_rows}->{$row_num} } );
        }

        for my $col_num ( $self->{_dim_colmin} .. $self->{_dim_colmax} ) {
            if ( my $col_ref = $self->{_table}->[$row_num]->[$col_num] ) {
                $self->_write_cell( $row_num, $col_num, $col_ref );
            }
        }

        $self->{_writer}->endTag( 'row' );
    }
    else {

        # Row attributes or comments only.
        $self->_write_empty_row( $row_num, undef,
            @{ $self->{_set_rows}->{$row_num} } );
    }

    # Reset table.
    $self->{_table} = [];

}


###############################################################################
#
# _calculate_spans()
#
# Calculate the "spans" attribute of the <row> tag. This is an XLSX
# optimisation and isn't strictly required. However, it makes comparing
# files easier.
#
# The span is the same for each block of 16 rows.
#
sub _calculate_spans {

    my $self = shift;

    my @spans;
    my $span_min;
    my $span_max;

    for my $row_num ( $self->{_dim_rowmin} .. $self->{_dim_rowmax} ) {

        # Calculate spans for cell data.
        if ( my $row_ref = $self->{_table}->[$row_num] ) {

            for my $col_num ( $self->{_dim_colmin} .. $self->{_dim_colmax} ) {
                if ( my $col_ref = $self->{_table}->[$row_num]->[$col_num] ) {

                    if ( !defined $span_min ) {
                        $span_min = $col_num;
                        $span_max = $col_num;
                    }
                    else {
                        $span_min = $col_num if $col_num < $span_min;
                        $span_max = $col_num if $col_num > $span_max;
                    }
                }
            }
        }

        # Calculate spans for comments.
        if ( defined $self->{_comments}->{$row_num} ) {

            for my $col_num ( $self->{_dim_colmin} .. $self->{_dim_colmax} ) {
                if ( defined $self->{_comments}->{$row_num}->{$col_num} ) {

                    if ( !defined $span_min ) {
                        $span_min = $col_num;
                        $span_max = $col_num;
                    }
                    else {
                        $span_min = $col_num if $col_num < $span_min;
                        $span_max = $col_num if $col_num > $span_max;
                    }
                }
            }
        }

        if ( ( ( $row_num + 1 ) % 16 == 0 )
            || $row_num == $self->{_dim_rowmax} )
        {
            my $span_index = int( $row_num / 16 );

            if ( defined $span_min ) {
                $span_min++;
                $span_max++;
                $spans[$span_index] = "$span_min:$span_max";
                $span_min = undef;
            }
        }
    }

    $self->{_row_spans} = \@spans;
}


###############################################################################
#
# _write_row()
#
# Write the <row> element.
#
sub _write_row {

    my $self      = shift;
    my $r         = shift;
    my $spans     = shift;
    my $height    = shift;
    my $format    = shift;
    my $hidden    = shift || 0;
    my $level     = shift || 0;
    my $collapsed = shift || 0;
    my $empty_row = shift || 0;
    my $xf_index  = 0;

    $height = 15 if !defined $height;

    my @attributes = ( 'r' => $r + 1 );

    # Get the format index.
    if ( ref( $format ) ) {
        $xf_index =  $format->get_xf_index();
    }

    push @attributes, ( 'spans'        => $spans )    if defined $spans;
    push @attributes, ( 's'            => $xf_index ) if $xf_index;
    push @attributes, ( 'customFormat' => 1 )         if $format;
    push @attributes, ( 'ht'           => $height )   if $height != 15;
    push @attributes, ( 'hidden'       => 1 )         if $hidden;
    push @attributes, ( 'customHeight' => 1 )         if $height != 15;
    push @attributes, ( 'outlineLevel' => $level )    if $level;
    push @attributes, ( 'collapsed'    => 1 )         if $collapsed;


    if ( $empty_row ) {
        $self->{_writer}->emptyTag( 'row', @attributes );
    }
    else {
        $self->{_writer}->startTag( 'row', @attributes );
    }
}


###############################################################################
#
# _write_empty_row()
#
# Write and empty <row> element, i.e., attributes only, no cell data.
#
sub _write_empty_row {

    my $self = shift;

    # Set the $empty_row parameter.
    $_[7] = 1;

    $self->_write_row( @_);
}


###############################################################################
#
# _write_cell()
#
# Write the <cell> element. This is the innermost loop so efficiency is
# important where possible. The basic methodology is that the data of every
# cell type is passed in as follows:
#
#      [ $row, $col, $aref]
#
# The aref, called $cell below, contains the following structure in all types:
#
#     [ $type, $token, $xf, @args ]
#
# Where $type:  represents the cell type, such as string, number, formula, etc.
#       $token: is the actual data for the string, number, formula, etc.
#       $xf:    is the XF format object.
#       @args:  additional args relevant to the specific data type.
#
sub _write_cell {

    my $self  = shift;
    my $row   = shift;
    my $col   = shift;
    my $cell  = shift;
    my $type  = $cell->[0];
    my $token = $cell->[1];
    my $xf    = $cell->[2];
    my $xf_index = 0;

    # Get the format index.
    if ( ref( $xf ) ) {
         $xf_index = $xf->get_xf_index();
    }

    my $range = xl_rowcol_to_cell( $row, $col );
    my @attributes = ( 'r' => $range );

    # Add the cell format index.
    if ( $xf_index ) {
        push @attributes, ( 's' => $xf_index );
    }
    elsif ( $self->{_set_rows}->{$row} && $self->{_set_rows}->{$row}->[1] ) {
        my $row_xf = $self->{_set_rows}->{$row}->[1];
        push @attributes, ( 's' => $row_xf->get_xf_index() );
    }
    elsif ( $self->{_col_formats}->{$col} ) {
        my $col_xf = $self->{_col_formats}->{$col};
        push @attributes, ( 's' => $col_xf->get_xf_index() );
    }


    # Write the various cell types.
    if ( $type eq 'n' ) {

        # Write a number.
        $self->{_writer}->startTag( 'c', @attributes );
        $self->_write_cell_value( $token );
        $self->{_writer}->endTag( 'c' );
    }
    elsif ( $type eq 's' ) {

        # Write a string.
        if ( $self->{_optimization} == 0 ) {
            push @attributes, ( 't' => 's' );
            $self->{_writer}->startTag( 'c', @attributes );
            $self->_write_cell_value( $token );
            $self->{_writer}->endTag( 'c' );
        }
        else {
            push @attributes, ( 't' => 'inlineStr' );
            $self->{_writer}->startTag( 'c', @attributes );
            $self->{_writer}->startTag( 'is' );

            my $string = $token;

            # Escape control characters. See SharedString.pm for details.
            $string =~ s/(_x[0-9a-fA-F]{4}_)/_x005F$1/g;
            $string =~ s/([\x00-\x08\x0B-\x1F])/sprintf "_x%04X_", ord($1)/eg;

            # Write any rich strings without further tags.
            if ( $string =~ m{^<r>} && $string =~ m{</r>$} ) {
                my $fh = $self->{_writer}->getOutput();

                local $\ = undef;    # Protect print from -l on commandline.
                print $fh $string;
            }
            else {
                my @t_attributes;

                # Add attribute to preserve leading or trailing whitespace.
                if ( $string =~ /^\s/ || $string =~ /\s$/ ) {
                    push @t_attributes, ( 'xml:space' => 'preserve' );
                }
                $self->{_writer}->dataElement( 't', $string, @t_attributes );
            }

            $self->{_writer}->endTag( 'is' );
            $self->{_writer}->endTag( 'c' );
        }
    }
    elsif ( $type eq 'f' ) {

        # Write a formula.
        $self->{_writer}->startTag( 'c', @attributes );
        $self->_write_cell_formula( $token );
        $self->_write_cell_value( $cell->[3] || 0 );
        $self->{_writer}->endTag( 'c' );
    }
    elsif ( $type eq 'a' ) {

        # Write an array formula.
        $self->{_writer}->startTag( 'c', @attributes );
        $self->_write_cell_array_formula( $token, $cell->[3] );
        $self->_write_cell_value( $cell->[4] );
        $self->{_writer}->endTag( 'c' );
    }
    elsif ( $type eq 'l' ) {
        my $link_type = $cell->[3];

        # Write the string part a hyperlink.
        push @attributes, ( 't' => 's' );

        $self->{_writer}->startTag( 'c', @attributes );
        $self->_write_cell_value( $token );
        $self->{_writer}->endTag( 'c' );

        if ( $link_type == 1 ) {

            # External link with rel file relationship.
            push @{ $self->{_hlink_refs} },
              [
                $link_type,              $row,       $col,
                ++$self->{_hlink_count}, $cell->[5], $cell->[6]
              ];

            push @{ $self->{_external_hyper_links} },
              [ '/hyperlink', $cell->[4], 'External' ];
        }
        elsif ( $link_type ) {

            # External link with rel file relationship.
            push @{ $self->{_hlink_refs} },
              [ $link_type, $row, $col, $cell->[4], $cell->[5], $cell->[6] ];
        }

    }
    elsif ( $type eq 'b' ) {

        # Write a empty cell.
        $self->{_writer}->emptyTag( 'c', @attributes );
    }
}


###############################################################################
#
# _write_cell_value()
#
# Write the cell value <v> element.
#
sub _write_cell_value {

    my $self = shift;
    my $value = defined $_[0] ? $_[0] : '';

    $self->{_writer}->dataElement( 'v', $value );
}


###############################################################################
#
# _write_cell_formula()
#
# Write the cell formula <f> element.
#
sub _write_cell_formula {

    my $self = shift;
    my $formula = defined $_[0] ? $_[0] : '';

    $self->{_writer}->dataElement( 'f', $formula );
}


###############################################################################
#
# _write_cell_array_formula()
#
# Write the cell array formula <f> element.
#
sub _write_cell_array_formula {

    my $self    = shift;
    my $formula = shift;
    my $range   = shift;

    my @attributes = ( 't' => 'array', 'ref' => $range );

    $self->{_writer}->dataElement( 'f', $formula, @attributes );
}


##############################################################################
#
# _write_sheet_calc_pr()
#
# Write the <sheetCalcPr> element for the worksheet calculation properties.
#
sub _write_sheet_calc_pr {

    my $self              = shift;
    my $full_calc_on_load = 1;

    my @attributes = ( 'fullCalcOnLoad' => $full_calc_on_load );

    $self->{_writer}->emptyTag( 'sheetCalcPr', @attributes );
}


###############################################################################
#
# _write_phonetic_pr()
#
# Write the <phoneticPr> element.
#
sub _write_phonetic_pr {

    my $self    = shift;
    my $font_id = 1;
    my $type    = 'noConversion';

    my @attributes = (
        'fontId' => $font_id,
        'type'   => $type,
    );

    $self->{_writer}->emptyTag( 'phoneticPr', @attributes );
}


###############################################################################
#
# _write_page_margins()
#
# Write the <pageMargins> element.
#
sub _write_page_margins {

    my $self = shift;

    my @attributes = (
        'left'   => $self->{_margin_left},
        'right'  => $self->{_margin_right},
        'top'    => $self->{_margin_top},
        'bottom' => $self->{_margin_bottom},
        'header' => $self->{_margin_header},
        'footer' => $self->{_margin_footer},
    );

    $self->{_writer}->emptyTag( 'pageMargins', @attributes );
}


###############################################################################
#
# _write_page_setup()
#
# Write the <pageSetup> element.
#
# The following is an example taken from Excel.
#
# <pageSetup
#     paperSize="9"
#     scale="110"
#     fitToWidth="2"
#     fitToHeight="2"
#     pageOrder="overThenDown"
#     orientation="portrait"
#     blackAndWhite="1"
#     draft="1"
#     horizontalDpi="200"
#     verticalDpi="200"
#     r:id="rId1"
# />
#
sub _write_page_setup {

    my $self       = shift;
    my @attributes = ();

    return unless $self->{_page_setup_changed};

    # Set paper size.
    if ( $self->{_paper_size} ) {
        push @attributes, ( 'paperSize' => $self->{_paper_size} );
    }

    # Set the print_scale
    if ( $self->{_print_scale} != 100 ) {
        push @attributes, ( 'scale' => $self->{_print_scale} );
    }

    # Set the "Fit to page" properties.
    if ( $self->{_fit_page} && $self->{_fit_width} != 1 ) {
        push @attributes, ( 'fitToWidth' => $self->{_fit_width} );
    }

    if ( $self->{_fit_page} && $self->{_fit_height} != 1 ) {
        push @attributes, ( 'fitToHeight' => $self->{_fit_height} );
    }

    # Set the page print direction.
    if ( $self->{_page_order} ) {
        push @attributes, ( 'pageOrder' => "overThenDown" );
    }

    # Set page orientation.
    if ( $self->{_orientation} == 0 ) {
        push @attributes, ( 'orientation' => 'landscape' );
    }
    else {
        push @attributes, ( 'orientation' => 'portrait' );
    }


    $self->{_writer}->emptyTag( 'pageSetup', @attributes );
}


##############################################################################
#
# _write_ext_lst()
#
# Write the <extLst> element.
#
sub _write_ext_lst {

    my $self = shift;

    $self->{_writer}->startTag( 'extLst' );
    $self->_write_ext();
    $self->{_writer}->endTag( 'extLst' );
}


###############################################################################
#
# _write_ext()
#
# Write the <ext> element.
#
sub _write_ext {

    my $self    = shift;
    my $xmlnsmx = 'http://schemas.microsoft.com/office/mac/excel/2008/main';
    my $uri     = 'http://schemas.microsoft.com/office/mac/excel/2008/main';

    my @attributes = (
        'xmlns:mx' => $xmlnsmx,
        'uri'      => $uri,
    );

    $self->{_writer}->startTag( 'ext', @attributes );
    $self->_write_mx_plv();
    $self->{_writer}->endTag( 'ext' );
}

###############################################################################
#
# _write_mx_plv()
#
# Write the <mx:PLV> element.
#
sub _write_mx_plv {

    my $self     = shift;
    my $mode     = 1;
    my $one_page = 0;
    my $w_scale  = 0;

    my @attributes = (
        'Mode'    => $mode,
        'OnePage' => $one_page,
        'WScale'  => $w_scale,
    );

    $self->{_writer}->emptyTag( 'mx:PLV', @attributes );
}


##############################################################################
#
# _write_merge_cells()
#
# Write the <mergeCells> element.
#
sub _write_merge_cells {

    my $self         = shift;
    my $merged_cells = $self->{_merge};
    my $count        = @$merged_cells;

    return unless $count;

    my @attributes = ( 'count' => $count );

    $self->{_writer}->startTag( 'mergeCells', @attributes );

    for my $merged_range ( @$merged_cells ) {

        # Write the mergeCell element.
        $self->_write_merge_cell( $merged_range );
    }

    $self->{_writer}->endTag( 'mergeCells' );
}


##############################################################################
#
# _write_merge_cell()
#
# Write the <mergeCell> element.
#
sub _write_merge_cell {

    my $self         = shift;
    my $merged_range = shift;
    my ( $row_min, $col_min, $row_max, $col_max ) = @$merged_range;


    # Convert the merge dimensions to a cell range.
    my $cell_1 = xl_rowcol_to_cell( $row_min, $col_min );
    my $cell_2 = xl_rowcol_to_cell( $row_max, $col_max );
    my $ref    = $cell_1 . ':' . $cell_2;

    my @attributes = ( 'ref' => $ref );

    $self->{_writer}->emptyTag( 'mergeCell', @attributes );
}


##############################################################################
#
# _write_print_options()
#
# Write the <printOptions> element.
#
sub _write_print_options {

    my $self       = shift;
    my @attributes = ();

    return unless $self->{_print_options_changed};

    # Set horizontal centering.
    if ( $self->{_hcenter} ) {
        push @attributes, ( 'horizontalCentered' => 1 );
    }

    # Set vertical centering.
    if ( $self->{_vcenter} ) {
        push @attributes, ( 'verticalCentered' => 1 );
    }

    # Enable row and column headers.
    if ( $self->{_print_headers} ) {
        push @attributes, ( 'headings' => 1 );
    }

    # Set printed gridlines.
    if ( $self->{_print_gridlines} ) {
        push @attributes, ( 'gridLines' => 1 );
    }


    $self->{_writer}->emptyTag( 'printOptions', @attributes );
}


##############################################################################
#
# _write_header_footer()
#
# Write the <headerFooter> element.
#
sub _write_header_footer {

    my $self = shift;

    return unless $self->{_header_footer_changed};

    $self->{_writer}->startTag( 'headerFooter' );
    $self->_write_odd_header() if $self->{_header};
    $self->_write_odd_footer() if $self->{_footer};
    $self->{_writer}->endTag( 'headerFooter' );
}


##############################################################################
#
# _write_odd_header()
#
# Write the <oddHeader> element.
#
sub _write_odd_header {

    my $self = shift;
    my $data = $self->{_header};

    $self->{_writer}->dataElement( 'oddHeader', $data );
}


##############################################################################
#
# _write_odd_footer()
#
# Write the <oddFooter> element.
#
sub _write_odd_footer {

    my $self = shift;
    my $data = $self->{_footer};

    $self->{_writer}->dataElement( 'oddFooter', $data );
}


##############################################################################
#
# _write_row_breaks()
#
# Write the <rowBreaks> element.
#
sub _write_row_breaks {

    my $self = shift;

    my @page_breaks = $self->_sort_pagebreaks( @{ $self->{_hbreaks} } );
    my $count       = scalar @page_breaks;

    return unless @page_breaks;

    my @attributes = (
        'count'            => $count,
        'manualBreakCount' => $count,
    );

    $self->{_writer}->startTag( 'rowBreaks', @attributes );

    for my $row_num ( @page_breaks ) {
        $self->_write_brk( $row_num, 16383 );
    }

    $self->{_writer}->endTag( 'rowBreaks' );
}


##############################################################################
#
# _write_col_breaks()
#
# Write the <colBreaks> element.
#
sub _write_col_breaks {

    my $self = shift;

    my @page_breaks = $self->_sort_pagebreaks( @{ $self->{_vbreaks} } );
    my $count       = scalar @page_breaks;

    return unless @page_breaks;

    my @attributes = (
        'count'            => $count,
        'manualBreakCount' => $count,
    );

    $self->{_writer}->startTag( 'colBreaks', @attributes );

    for my $col_num ( @page_breaks ) {
        $self->_write_brk( $col_num, 1048575 );
    }

    $self->{_writer}->endTag( 'colBreaks' );
}


##############################################################################
#
# _write_brk()
#
# Write the <brk> element.
#
sub _write_brk {

    my $self = shift;
    my $id   = shift;
    my $max  = shift;
    my $man  = 1;

    my @attributes = (
        'id'  => $id,
        'max' => $max,
        'man' => $man,
    );

    $self->{_writer}->emptyTag( 'brk', @attributes );
}


##############################################################################
#
# _write_auto_filter()
#
# Write the <autoFilter> element.
#
sub _write_auto_filter {

    my $self = shift;
    my $ref  = $self->{_autofilter_ref};

    return unless $ref;

    my @attributes = ( 'ref' => $ref );

    if ( $self->{_filter_on} ) {

        # Autofilter defined active filters.
        $self->{_writer}->startTag( 'autoFilter', @attributes );

        $self->_write_autofilters();

        $self->{_writer}->endTag( 'autoFilter' );

    }
    else {

        # Autofilter defined without active filters.
        $self->{_writer}->emptyTag( 'autoFilter', @attributes );
    }

}


###############################################################################
#
# _write_autofilters()
#
# Function to iterate through the columns that form part of an autofilter
# range and write the appropriate filters.
#
sub _write_autofilters {

    my $self = shift;

    my ( $col1, $col2 ) = @{ $self->{_filter_range} };

    for my $col ( $col1 .. $col2 ) {

        # Skip if column doesn't have an active filter.
        next unless $self->{_filter_cols}->{$col};

        # Retrieve the filter tokens and write the autofilter records.
        my @tokens = @{ $self->{_filter_cols}->{$col} };
        my $type   = $self->{_filter_type}->{$col};

        $self->_write_filter_column( $col, $type, \@tokens );
    }
}


##############################################################################
#
# _write_filter_column()
#
# Write the <filterColumn> element.
#
sub _write_filter_column {

    my $self    = shift;
    my $col_id  = shift;
    my $type    = shift;
    my $filters = shift;

    my @attributes = ( 'colId' => $col_id );

    $self->{_writer}->startTag( 'filterColumn', @attributes );


    if ( $type == 1 ) {

        # Type == 1 is the new XLSX style filter.
        $self->_write_filters( @$filters );

    }
    else {

        # Type == 0 is the classic "custom" filter.
        $self->_write_custom_filters( @$filters );
    }

    $self->{_writer}->endTag( 'filterColumn' );
}


##############################################################################
#
# _write_filters()
#
# Write the <filters> element.
#
sub _write_filters {

    my $self    = shift;
    my @filters = @_;

    if ( @filters == 1 && $filters[0] eq 'blanks' ) {

        # Special case for blank cells only.
        $self->{_writer}->emptyTag( 'filters', 'blank' => 1 );
    }
    else {

        # General case.
        $self->{_writer}->startTag( 'filters' );

        for my $filter ( @filters ) {
            $self->_write_filter( $filter );
        }

        $self->{_writer}->endTag( 'filters' );
    }
}


##############################################################################
#
# _write_filter()
#
# Write the <filter> element.
#
sub _write_filter {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_writer}->emptyTag( 'filter', @attributes );
}


##############################################################################
#
# _write_custom_filters()
#
# Write the <customFilters> element.
#
sub _write_custom_filters {

    my $self   = shift;
    my @tokens = @_;

    if ( @tokens == 2 ) {

        # One filter expression only.
        $self->{_writer}->startTag( 'customFilters' );
        $self->_write_custom_filter( @tokens );
        $self->{_writer}->endTag( 'customFilters' );

    }
    else {

        # Two filter expressions.

        my @attributes;

        # Check if the "join" operand is "and" or "or".
        if ( $tokens[2] == 0 ) {
            @attributes = ( 'and' => 1 );
        }
        else {
            @attributes = ( 'and' => 0 );
        }

        # Write the two custom filters.
        $self->{_writer}->startTag( 'customFilters', @attributes );
        $self->_write_custom_filter( $tokens[0], $tokens[1] );
        $self->_write_custom_filter( $tokens[3], $tokens[4] );
        $self->{_writer}->endTag( 'customFilters' );
    }
}


##############################################################################
#
# _write_custom_filter()
#
# Write the <customFilter> element.
#
sub _write_custom_filter {

    my $self       = shift;
    my $operator   = shift;
    my $val        = shift;
    my @attributes = ();

    my %operators = (
        1  => 'lessThan',
        2  => 'equal',
        3  => 'lessThanOrEqual',
        4  => 'greaterThan',
        5  => 'notEqual',
        6  => 'greaterThanOrEqual',
        22 => 'equal',
    );


    # Convert the operator from a number to a descriptive string.
    if ( defined $operators{$operator} ) {
        $operator = $operators{$operator};
    }
    else {
        croak "Unknown operator = $operator\n";
    }

    # The 'equal' operator is the default attribute and isn't stored.
    push @attributes, ( 'operator' => $operator ) unless $operator eq 'equal';
    push @attributes, ( 'val' => $val );

    $self->{_writer}->emptyTag( 'customFilter', @attributes );
}


##############################################################################
#
# _write_hyperlinks()
#
# Write the <hyperlinks> element. The attributes are different for internal
# and external links.
#
sub _write_hyperlinks {

    my $self       = shift;
    my @hlink_refs = @{ $self->{_hlink_refs} };

    return unless @hlink_refs;

    $self->{_writer}->startTag( 'hyperlinks' );

    for my $aref ( @hlink_refs ) {
        my ( $type, @args ) = @$aref;

        if ( $type == 1 ) {
            $self->_write_hyperlink_external( @args );
        }
        elsif ( $type == 2 ) {
            $self->_write_hyperlink_internal( @args );
        }
    }

    $self->{_writer}->endTag( 'hyperlinks' );
}


##############################################################################
#
# _write_hyperlink_external()
#
# Write the <hyperlink> element for external links.
#
sub _write_hyperlink_external {

    my $self     = shift;
    my $row      = shift;
    my $col      = shift;
    my $id       = shift;
    my $location = shift;
    my $tooltip  = shift;

    my $ref = xl_rowcol_to_cell( $row, $col );
    my $r_id = 'rId' . $id;

    my @attributes = (
        'ref'  => $ref,
        'r:id' => $r_id,
    );

    push @attributes, ( 'location' => $location ) if defined $location;
    push @attributes, ( 'tooltip'  => $tooltip )  if defined $tooltip;

    $self->{_writer}->emptyTag( 'hyperlink', @attributes );
}


##############################################################################
#
# _write_hyperlink_internal()
#
# Write the <hyperlink> element for internal links.
#
sub _write_hyperlink_internal {

    my $self     = shift;
    my $row      = shift;
    my $col      = shift;
    my $location = shift;
    my $display  = shift;
    my $tooltip  = shift;

    my $ref = xl_rowcol_to_cell( $row, $col );

    my @attributes = ( 'ref' => $ref, 'location' => $location );

    push @attributes, ( 'tooltip' => $tooltip ) if defined $tooltip;
    push @attributes, ( 'display' => $display );

    $self->{_writer}->emptyTag( 'hyperlink', @attributes );
}


##############################################################################
#
# _write_panes()
#
# Write the frozen or split <pane> elements.
#
sub _write_panes {

    my $self  = shift;
    my @panes = @{ $self->{_panes} };

    return unless @panes;

    if ( $panes[4] == 2 ) {
        $self->_write_split_panes( @panes );
    }
    else {
        $self->_write_freeze_panes( @panes );
    }
}


##############################################################################
#
# _write_freeze_panes()
#
# Write the <pane> element for freeze panes.
#
sub _write_freeze_panes {

    my $self = shift;
    my @attributes;

    my ( $row, $col, $top_row, $left_col, $type ) = @_;

    my $y_split       = $row;
    my $x_split       = $col;
    my $top_left_cell = xl_rowcol_to_cell( $top_row, $left_col );
    my $active_pane;
    my $state;
    my $active_cell;
    my $sqref;

    # Move user cell selection to the panes.
    if ( @{ $self->{_selections} } ) {
        ( undef, $active_cell, $sqref ) = @{ $self->{_selections}->[0] };
        $self->{_selections} = [];
    }

    # Set the active pane.
    if ( $row && $col ) {
        $active_pane = 'bottomRight';

        my $row_cell = xl_rowcol_to_cell( $row, 0 );
        my $col_cell = xl_rowcol_to_cell( 0,    $col );

        push @{ $self->{_selections} },
          (
            [ 'topRight',    $col_cell,    $col_cell ],
            [ 'bottomLeft',  $row_cell,    $row_cell ],
            [ 'bottomRight', $active_cell, $sqref ]
          );
    }
    elsif ( $col ) {
        $active_pane = 'topRight';
        push @{ $self->{_selections} }, [ 'topRight', $active_cell, $sqref ];
    }
    else {
        $active_pane = 'bottomLeft';
        push @{ $self->{_selections} }, [ 'bottomLeft', $active_cell, $sqref ];
    }

    # Set the pane type.
    if ( $type == 0 ) {
        $state = 'frozen';
    }
    elsif ( $type == 1 ) {
        $state = 'frozenSplit';
    }
    else {
        $state = 'split';
    }


    push @attributes, ( 'xSplit' => $x_split ) if $x_split;
    push @attributes, ( 'ySplit' => $y_split ) if $y_split;

    push @attributes, ( 'topLeftCell' => $top_left_cell );
    push @attributes, ( 'activePane'  => $active_pane );
    push @attributes, ( 'state'       => $state );


    $self->{_writer}->emptyTag( 'pane', @attributes );
}


##############################################################################
#
# _write_split_panes()
#
# Write the <pane> element for split panes.
#
# See also, implementers note for split_panes().
#
sub _write_split_panes {

    my $self = shift;
    my @attributes;
    my $y_split;
    my $x_split;
    my $has_selection = 0;
    my $active_pane;
    my $active_cell;
    my $sqref;

    my ( $row, $col, $top_row, $left_col, $type ) = @_;
    $y_split = $row;
    $x_split = $col;

    # Move user cell selection to the panes.
    if ( @{ $self->{_selections} } ) {
        ( undef, $active_cell, $sqref ) = @{ $self->{_selections}->[0] };
        $self->{_selections} = [];
        $has_selection = 1;
    }

    # Convert the row and col to 1/20 twip units with padding.
    $y_split = int( 20 * $y_split + 300 ) if $y_split;
    $x_split = $self->_calculate_x_split_width( $x_split ) if $x_split;

    # For non-explicit topLeft definitions, estimate the cell offset based
    # on the pixels dimensions. This is only a workaround and doesn't take
    # adjusted cell dimensions into account.
    if ( $top_row == $row && $left_col == $col ) {
        $top_row  = int( 0.5 + ( $y_split - 300 ) / 20 / 15 );
        $left_col = int( 0.5 + ( $x_split - 390 ) / 20 / 3 * 4 / 64 );
    }

    my $top_left_cell = xl_rowcol_to_cell( $top_row, $left_col );

    # If there is no selection set the active cell to the top left cell.
    if ( !$has_selection ) {
        $active_cell = $top_left_cell;
        $sqref       = $top_left_cell;
    }

    # Set the Cell selections.
    if ( $row && $col ) {
        $active_pane = 'bottomRight';

        my $row_cell = xl_rowcol_to_cell( $top_row, 0 );
        my $col_cell = xl_rowcol_to_cell( 0,        $left_col );

        push @{ $self->{_selections} },
          (
            [ 'topRight',    $col_cell,    $col_cell ],
            [ 'bottomLeft',  $row_cell,    $row_cell ],
            [ 'bottomRight', $active_cell, $sqref ]
          );
    }
    elsif ( $col ) {
        $active_pane = 'topRight';
        push @{ $self->{_selections} }, [ 'topRight', $active_cell, $sqref ];
    }
    else {
        $active_pane = 'bottomLeft';
        push @{ $self->{_selections} }, [ 'bottomLeft', $active_cell, $sqref ];
    }

    push @attributes, ( 'xSplit' => $x_split ) if $x_split;
    push @attributes, ( 'ySplit' => $y_split ) if $y_split;
    push @attributes, ( 'topLeftCell' => $top_left_cell );
    push @attributes, ( 'activePane' => $active_pane ) if $has_selection;

    $self->{_writer}->emptyTag( 'pane', @attributes );
}


##############################################################################
#
# _calculate_x_split_width()
#
# Convert column width from user units to pane split width.
#
sub _calculate_x_split_width {

    my $self  = shift;
    my $width = shift;

    my $max_digit_width = 7;    # For Calabri 11.
    my $padding         = 5;
    my $pixels;

    # Convert to pixels.
    if ( $width < 1 ) {
        $pixels = int( $width * 12 + 0.5 );
    }
    else {
        $pixels = int( $width * $max_digit_width + 0.5 ) + $padding;
    }

    # Convert to points.
    my $points = $pixels * 3 / 4;

    # Convert to twips (twentieths of a point).
    my $twips = $points * 20;

    # Add offset/padding.
    $width = $twips + 390;

    return $width;
}


##############################################################################
#
# _write_tab_color()
#
# Write the <tabColor> element.
#
sub _write_tab_color {

    my $self        = shift;
    my $color_index = $self->{_tab_color};

    return unless $color_index;

    my $rgb = $self->_get_palette_color( $color_index );

    my @attributes = ( 'rgb' => $rgb );

    $self->{_writer}->emptyTag( 'tabColor', @attributes );
}


##############################################################################
#
# _write_outline_pr()
#
# Write the <outlinePr> element.
#
sub _write_outline_pr {

    my $self        = shift;
    my @attributes = ();

    return unless $self->{_outline_changed};

    push @attributes, ( "applyStyles"  => 1 ) if $self->{_outline_style};
    push @attributes, ( "summaryBelow" => 0 ) if !$self->{_outline_below};
    push @attributes, ( "summaryRight" => 0 ) if !$self->{_outline_right};
    push @attributes, ( "showOutlineSymbols" => 0 ) if !$self->{_outline_on};

    $self->{_writer}->emptyTag( 'outlinePr', @attributes );
}


##############################################################################
#
# _write_sheet_protection()
#
# Write the <sheetProtection> element.
#
sub _write_sheet_protection {

    my $self = shift;
    my @attributes;

    return unless $self->{_protect};

    my %arg = %{ $self->{_protect} };

    push @attributes, ( "password" => $arg{password} ) if $arg{password};
    push @attributes, ( "sheet"            => 1 ) if $arg{sheet};
    push @attributes, ( "content"          => 1 ) if $arg{content};
    push @attributes, ( "objects"          => 1 ) if !$arg{objects};
    push @attributes, ( "scenarios"        => 1 ) if !$arg{scenarios};
    push @attributes, ( "formatCells"      => 0 ) if $arg{format_cells};
    push @attributes, ( "formatColumns"    => 0 ) if $arg{format_columns};
    push @attributes, ( "formatRows"       => 0 ) if $arg{format_rows};
    push @attributes, ( "insertColumns"    => 0 ) if $arg{insert_columns};
    push @attributes, ( "insertRows"       => 0 ) if $arg{insert_rows};
    push @attributes, ( "insertHyperlinks" => 0 ) if $arg{insert_hyperlinks};
    push @attributes, ( "deleteColumns"    => 0 ) if $arg{delete_columns};
    push @attributes, ( "deleteRows"       => 0 ) if $arg{delete_rows};

    push @attributes, ( "selectLockedCells" => 1 )
      if !$arg{select_locked_cells};

    push @attributes, ( "sort"        => 0 ) if $arg{sort};
    push @attributes, ( "autoFilter"  => 0 ) if $arg{autofilter};
    push @attributes, ( "pivotTables" => 0 ) if $arg{pivot_tables};

    push @attributes, ( "selectUnlockedCells" => 1 )
      if !$arg{select_unlocked_cells};


    $self->{_writer}->emptyTag( 'sheetProtection', @attributes );
}


##############################################################################
#
# _write_drawings()
#
# Write the <drawing> elements.
#
sub _write_drawings {

    my $self = shift;

    return unless $self->{_drawing};

    $self->_write_drawing( $self->{_hlink_count} + 1 );
}


##############################################################################
#
# _write_drawing()
#
# Write the <drawing> element.
#
sub _write_drawing {

    my $self = shift;
    my $id   = shift;
    my $r_id = 'rId' . $id;

    my @attributes = ( 'r:id' => $r_id );

    $self->{_writer}->emptyTag( 'drawing', @attributes );
}


##############################################################################
#
# _write_legacy_drawing()
#
# Write the <legacyDrawing> element.
#
sub _write_legacy_drawing {

    my $self = shift;
    my $id;

    return unless $self->{_has_comments};

    # Increment the relationship id for any drawings or comments.
    $id = $self->{_hlink_count} + 1;
    $id++ if $self->{_drawing};


    my @attributes = ( 'r:id' => 'rId' . $id );

    $self->{_writer}->emptyTag( 'legacyDrawing', @attributes );
}


#
# Note, the following font methods are, more or less, duplicated from the
# Excel::Writer::XLSX::Package::Styles class. I will look at implementing
# this is a cleaner encapsulated mode at a later stage.
#


##############################################################################
#
# _write_font()
#
# Write the <font> element.
#
sub _write_font {

    my $self   = shift;
    my $format = shift;

    $self->{_rstring}->startTag( 'rPr' );

    $self->{_rstring}->emptyTag( 'b' )       if $format->{_bold};
    $self->{_rstring}->emptyTag( 'i' )       if $format->{_italic};
    $self->{_rstring}->emptyTag( 'strike' )  if $format->{_font_strikeout};
    $self->{_rstring}->emptyTag( 'outline' ) if $format->{_font_outline};
    $self->{_rstring}->emptyTag( 'shadow' )  if $format->{_font_shadow};

    # Handle the underline variants.
    $self->_write_underline( $format->{_underline} ) if $format->{_underline};

    $self->_write_vert_align( 'superscript' ) if $format->{_font_script} == 1;
    $self->_write_vert_align( 'subscript' )   if $format->{_font_script} == 2;

    $self->{_rstring}->emptyTag( 'sz', 'val', $format->{_size} );

    if ( my $theme = $format->{_theme} ) {
        $self->_write_rstring_color( 'theme' => $theme );
    }
    elsif ( my $color = $format->{_color} ) {
        $color = $self->_get_palette_color( $color );

        $self->_write_rstring_color( 'rgb' => $color );
    }
    else {
        $self->_write_rstring_color( 'theme' => 1 );
    }

    $self->{_rstring}->emptyTag( 'rFont',  'val', $format->{_font} );
    $self->{_rstring}->emptyTag( 'family', 'val', $format->{_font_family} );

    if ( $format->{_font} eq 'Calibri' && !$format->{_hyperlink} ) {
        $self->{_rstring}->emptyTag( 'scheme', 'val', $format->{_font_scheme} );
    }

    $self->{_rstring}->endTag( 'rPr' );
}


###############################################################################
#
# _write_underline()
#
# Write the underline font element.
#
sub _write_underline {

    my $self      = shift;
    my $underline = shift;
    my @attributes;

    # Handle the underline variants.
    if ( $underline == 2 ) {
        @attributes = ( val => 'double' );
    }
    elsif ( $underline == 33 ) {
        @attributes = ( val => 'singleAccounting' );
    }
    elsif ( $underline == 34 ) {
        @attributes = ( val => 'doubleAccounting' );
    }
    else {
        @attributes = ();    # Default to single underline.
    }

    $self->{_rstring}->emptyTag( 'u', @attributes );

}


##############################################################################
#
# _write_vert_align()
#
# Write the <vertAlign> font sub-element.
#
sub _write_vert_align {

    my $self = shift;
    my $val  = shift;

    my @attributes = ( 'val' => $val );

    $self->{_rstring}->emptyTag( 'vertAlign', @attributes );
}


##############################################################################
#
# _write_rstring_color()
#
# Write the <color> element.
#
sub _write_rstring_color {

    my $self  = shift;
    my $name  = shift;
    my $value = shift;

    my @attributes = ( $name => $value );

    $self->{_rstring}->emptyTag( 'color', @attributes );
}


#
# End font duplication code.
#


##############################################################################
#
# _write_data_validations()
#
# Write the <dataValidations> element.
#
sub _write_data_validations {

    my $self        = shift;
    my @validations = @{ $self->{_validations} };
    my $count       = @validations;

    return unless $count;

    my @attributes = ( 'count' => $count );

    $self->{_writer}->startTag( 'dataValidations', @attributes );

    for my $validation ( @validations ) {

        # Write the dataValidation element.
        $self->_write_data_validation( $validation );
    }

    $self->{_writer}->endTag( 'dataValidations' );
}


##############################################################################
#
# _write_data_validation()
#
# Write the <dataValidation> element.
#
sub _write_data_validation {

    my $self       = shift;
    my $param      = shift;
    my $sqref      = '';
    my @attributes = ();


    # Set the cell range(s) for the data validation.
    for my $cells ( @{ $param->{cells} } ) {

        # Add a space between multiple cell ranges.
        $sqref .= ' ' if $sqref ne '';

        my ( $row_first, $col_first, $row_last, $col_last ) = @$cells;

        # Swap last row/col for first row/col as necessary
        if ( $row_first > $row_last ) {
            ( $row_first, $row_last ) = ( $row_last, $row_first );
        }

        if ( $col_first > $col_last ) {
            ( $col_first, $col_last ) = ( $col_last, $col_first );
        }

        # If the first and last cell are the same write a single cell.
        if ( ( $row_first == $row_last ) && ( $col_first == $col_last ) ) {
            $sqref .= xl_rowcol_to_cell( $row_first, $col_first );
        }
        else {
            $sqref .= xl_range( $row_first, $row_last, $col_first, $col_last );
        }
    }


    push @attributes, ( 'type' => $param->{validate} );

    if ( $param->{criteria} ne 'between' ) {
        push @attributes, ( 'operator' => $param->{criteria} );
    }

    if ( $param->{error_type} ) {
        push @attributes, ( 'errorStyle' => 'warning' )
          if $param->{error_type} == 1;
        push @attributes, ( 'errorStyle' => 'information' )
          if $param->{error_type} == 2;
    }

    push @attributes, ( 'allowBlank'       => 1 ) if $param->{ignore_blank};
    push @attributes, ( 'showDropDown'     => 1 ) if !$param->{dropdown};
    push @attributes, ( 'showInputMessage' => 1 ) if $param->{show_input};
    push @attributes, ( 'showErrorMessage' => 1 ) if $param->{show_error};

    push @attributes, ( 'errorTitle' => $param->{error_title} )
      if $param->{error_title};

    push @attributes, ( 'error' => $param->{error_message} )
      if $param->{error_message};

    push @attributes, ( 'promptTitle' => $param->{input_title} )
      if $param->{input_title};

    push @attributes, ( 'prompt' => $param->{input_message} )
      if $param->{input_message};

    push @attributes, ( 'sqref' => $sqref );

    $self->{_writer}->startTag( 'dataValidation', @attributes );

    # Write the formula1 element.
    $self->_write_formula_1( $param->{value} );

    # Write the formula2 element.
    $self->_write_formula_2( $param->{maximum} ) if defined $param->{maximum};

    $self->{_writer}->endTag( 'dataValidation' );
}


##############################################################################
#
# _write_formula_1()
#
# Write the <formula1> element.
#
sub _write_formula_1 {

    my $self    = shift;
    my $formula = shift;

    # Convert a list array ref into a comma separated string.
    if (ref $formula eq 'ARRAY') {
        $formula   = join ',', @$formula;
        $formula   = qq("$formula");
    }

    $formula =~ s/^=//;    # Remove formula symbol.

    $self->{_writer}->dataElement( 'formula1', $formula );
}


##############################################################################
#
# _write_formula_2()
#
# Write the <formula2> element.
#
sub _write_formula_2 {

    my $self    = shift;
    my $formula = shift;

    $formula =~ s/^=//;    # Remove formula symbol.

    $self->{_writer}->dataElement( 'formula2', $formula );
}


##############################################################################
#
# _write_conditional_formats()
#
# Write the Worksheet conditional formats.
#
sub _write_conditional_formats {

    my $self     = shift;
    my @ranges   = sort keys %{ $self->{_cond_formats} };

    return unless scalar @ranges;

    for my $range ( @ranges ) {
        $self->_write_conditional_formatting( $range,
            $self->{_cond_formats}->{$range} );
    }
}


##############################################################################
#
# _write_conditional_formatting()
#
# Write the <conditionalFormatting> element.
#
sub _write_conditional_formatting {

    my $self   = shift;
    my $range  = shift;
    my $params = shift;

    my @attributes = ( 'sqref' => $range );

    $self->{_writer}->startTag( 'conditionalFormatting', @attributes );

    for my $param ( @$params ) {

        # Write the cfRule element.
        $self->_write_cf_rule( $param );
    }

    $self->{_writer}->endTag( 'conditionalFormatting' );
}

##############################################################################
#
# _write_cf_rule()
#
# Write the <cfRule> element.
#
sub _write_cf_rule {

    my $self  = shift;
    my $param = shift;

    my @attributes = ( 'type' => $param->{type} );

    push @attributes, ( 'dxfId' => $param->{format} )
      if defined $param->{format};

    push @attributes, ( 'priority' => $param->{priority} );

    if ( $param->{type} eq 'cellIs' ) {
        push @attributes, ( 'operator' => $param->{criteria} );

        $self->{_writer}->startTag( 'cfRule', @attributes );

        if ( defined $param->{minimum} && defined $param->{maximum} ) {
            $self->_write_formula( $param->{minimum} );
            $self->_write_formula( $param->{maximum} );
        }
        else {
            $self->_write_formula( $param->{value} );
        }

        $self->{_writer}->endTag( 'cfRule' );
    }
    elsif ( $param->{type} eq 'aboveAverage' ) {
        if ( $param->{criteria} =~ /below/ ) {
            push @attributes, ( 'aboveAverage' => 0 );
        }

        if ( $param->{criteria} =~ /equal/ ) {
            push @attributes, ( 'equalAverage' => 1 );
        }

        if ( $param->{criteria} =~ /([123]) std dev/ ) {
            push @attributes, ( 'stdDev' => $1 );
        }

        $self->{_writer}->emptyTag( 'cfRule', @attributes );
    }
    elsif ( $param->{type} eq 'top10' ) {
        if ( defined $param->{criteria} && $param->{criteria} eq '%' ) {
            push @attributes, ( 'percent' => 1 );
        }

        if ( $param->{direction} ) {
            push @attributes, ( 'bottom' => 1 );
        }

        my $rank = $param->{value} || 10;
        push @attributes, ( 'rank' => $rank );

        $self->{_writer}->emptyTag( 'cfRule', @attributes );
    }
    elsif ( $param->{type} eq 'duplicateValues' ) {
        $self->{_writer}->emptyTag( 'cfRule', @attributes );
    }
    elsif ( $param->{type} eq 'uniqueValues' ) {
        $self->{_writer}->emptyTag( 'cfRule', @attributes );
    }
    elsif ($param->{type} eq 'containsText'
        || $param->{type} eq 'notContainsText'
        || $param->{type} eq 'beginsWith'
        || $param->{type} eq 'endsWith' )
    {
        push @attributes, ( 'operator' => $param->{criteria} );
        push @attributes, ( 'text'     => $param->{value} );

        $self->{_writer}->startTag( 'cfRule', @attributes );
        $self->_write_formula( $param->{formula} );
        $self->{_writer}->endTag( 'cfRule' );
    }
    elsif ( $param->{type} eq 'timePeriod' ) {
        push @attributes, ( 'timePeriod' => $param->{criteria} );

        $self->{_writer}->startTag( 'cfRule', @attributes );
        $self->_write_formula( $param->{formula} );
        $self->{_writer}->endTag( 'cfRule' );
    }
    elsif ($param->{type} eq 'containsBlanks'
        || $param->{type} eq 'notContainsBlanks'
        || $param->{type} eq 'containsErrors'
        || $param->{type} eq 'notContainsErrors' )
    {
        $self->{_writer}->startTag( 'cfRule', @attributes );
        $self->_write_formula( $param->{formula} );
        $self->{_writer}->endTag( 'cfRule' );
    }
    elsif ( $param->{type} eq 'colorScale' ) {

        $self->{_writer}->startTag( 'cfRule', @attributes );
        $self->_write_color_scale( $param );
        $self->{_writer}->endTag( 'cfRule' );
    }
    elsif ( $param->{type} eq 'dataBar' ) {

        $self->{_writer}->startTag( 'cfRule', @attributes );
        $self->_write_data_bar( $param );
        $self->{_writer}->endTag( 'cfRule' );
    }
    elsif ( $param->{type} eq 'expression' ) {

        $self->{_writer}->startTag( 'cfRule', @attributes );
        $self->_write_formula( $param->{criteria} );
        $self->{_writer}->endTag( 'cfRule' );
    }
}


##############################################################################
#
# _write_formula()
#
# Write the <formula> element.
#
sub _write_formula {

    my $self = shift;
    my $data = shift;

    # Remove equality from formula.
    $data =~ s/^=//;

    $self->{_writer}->dataElement( 'formula', $data );
}


##############################################################################
#
# _write_color_scale()
#
# Write the <colorScale> element.
#
sub _write_color_scale {

    my $self  = shift;
    my $param = shift;

    $self->{_writer}->startTag( 'colorScale' );

    $self->_write_cfvo( $param->{min_type}, $param->{min_value} );

    if ( defined $param->{mid_type} ) {
        $self->_write_cfvo( $param->{mid_type}, $param->{mid_value} );
    }

    $self->_write_cfvo( $param->{max_type}, $param->{max_value} );

    $self->_write_color( 'rgb' => $param->{min_color} );

    if ( defined $param->{mid_color} ) {
        $self->_write_color( 'rgb' => $param->{mid_color} );
    }

    $self->_write_color( 'rgb' => $param->{max_color} );

    $self->{_writer}->endTag( 'colorScale' );
}


##############################################################################
#
# _write_data_bar()
#
# Write the <dataBar> element.
#
sub _write_data_bar {

    my $self  = shift;
    my $param = shift;

    $self->{_writer}->startTag( 'dataBar' );

    $self->_write_cfvo( $param->{min_type}, $param->{min_value} );
    $self->_write_cfvo( $param->{max_type}, $param->{max_value} );

    $self->_write_color( 'rgb' => $param->{bar_color} );

    $self->{_writer}->endTag( 'dataBar' );
}


##############################################################################
#
# _write_cfvo()
#
# Write the <cfvo> element.
#
sub _write_cfvo {

    my $self = shift;
    my $type = shift;
    my $val  = shift;

    my @attributes = (
        'type' => $type,
        'val'  => $val
    );

    $self->{_writer}->emptyTag( 'cfvo', @attributes );
}



##############################################################################
#
# _write_color()
#
# Write the <color> element.
#
sub _write_color {

    my $self  = shift;
    my $name  = shift;
    my $value = shift;

    my @attributes = ( $name => $value );

    $self->{_writer}->emptyTag( 'color', @attributes );
}


1;


__END__


=head1 NAME

Worksheet - A class for writing Excel Worksheets.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

� MM-MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

