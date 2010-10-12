package Excel::Writer::XLSX::Worksheet;

###############################################################################
#
# Worksheet - A writer class for Excel Worksheets.
#
#
# Used in conjunction with Excel::Writer::XLSX
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
use Excel::Writer::XLSX::Utility qw(xl_cell_to_rowcol xl_rowcol_to_cell);

our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
our $VERSION = '0.02';


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

    $self->{_name}        = $_[0];
    $self->{_index}       = $_[1];
    $self->{_activesheet} = $_[2];
    $self->{_firstsheet}  = $_[3];
    $self->{_str_total}   = $_[4];
    $self->{_str_unique}  = $_[5];
    $self->{_str_table}   = $_[6];
    $self->{_1904}        = $_[7];

    $self->{_ext_sheets}  = [];
    $self->{_fileclosed}  = 0;

    $self->{_xls_rowmax}  = $rowmax;
    $self->{_xls_colmax}  = $colmax;
    $self->{_xls_strmax}  = $strmax;
    $self->{_dim_rowmin}  = undef;
    $self->{_dim_rowmax}  = undef;
    $self->{_dim_colmin}  = undef;
    $self->{_dim_colmax}  = undef;

    $self->{_colinfo}     = [];
    $self->{_selection}   = [ 0, 0 ];
    $self->{_hidden}      = 0;
    $self->{_active}      = 0;
    $self->{_tab_color}   = 0;

    $self->{_panes}       = [];
    $self->{_active_pane} = 3;
    $self->{_frozen}      = 0;
    $self->{_selected}    = 0;

    $self->{_paper_size}    = 0x0;
    $self->{_orientation}   = 0x1;
    $self->{_header}        = '';
    $self->{_footer}        = '';
    $self->{_hcenter}       = 0;
    $self->{_vcenter}       = 0;
    $self->{_margin_header}   = 0.50;
    $self->{_margin_footer}   = 0.50;
    $self->{_margin_left}   = 0.75;
    $self->{_margin_right}  = 0.75;
    $self->{_margin_top}    = 1.00;
    $self->{_margin_bottom} = 1.00;

    $self->{_repeat_rows} = '';
    $self->{_repeat_cols} = '';

    $self->{_print_gridlines}  = 0;
    $self->{_screen_gridlines} = 1;
    $self->{_print_headers}    = 0;

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

    $self->{_zoom}        = 100;
    $self->{_print_scale} = 100;

    $self->{_leading_zeros} = 0;

    $self->{_outline_row_level} = 0;
    $self->{_outline_style}     = 0;
    $self->{_outline_below}     = 1;
    $self->{_outline_right}     = 1;
    $self->{_outline_on}        = 1;

    $self->{_names} = {};

    $self->{_write_match} = [];

    $self->{prev_col} = -1;

    $self->{_table}   = [];
    $self->{_merge}   = {};
    $self->{_comment} = {};

    $self->{_autofilter}   = '';
    $self->{_filter_on}    = 0;
    $self->{_filter_range} = [];
    $self->{_filter_cols}  = {};

    $self->{_col_sizes}   = {};
    $self->{_row_sizes}   = {};
    $self->{_col_formats} = {};
    $self->{_row_formats} = {};


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

    $self->_write_xml_declaration();

    # Write the root worksheet element.
    $self->_write_worksheet();

    # Write the worksheet properties.
    #$self->_write_sheet_pr();

    # Write the worksheet dimensions.
    $self->_write_dimension();

    # Write the sheet view properties.
    $self->_write_sheet_views();

    # Write the sheet format properties.
    $self->_write_sheet_format_pr();

    # Write the sheet column info.
    $self->_write_cols();

    # Write the worksheet data such as rows columns and cells.
    $self->_write_sheet_data();

    # Write the worksheet calculation properties.
    #$self->_write_sheet_calc_pr();

    # Write the worksheet phonetic properties.
    #$self->_write_phonetic_pr();

    # Write the worksheet page_margins.
    $self->_write_page_margins();

    # Write the worksheet page setup.
    #$self->_write_page_setup();

    # Write the worksheet extension storage.
    #$self->_write_ext_lst();

    # Close the worksheet tag.
    $self->{_writer}->endTag( 'worksheet' );

    # Close the XML::Writer object and filehandle.
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
# protect($password)
#
# Set the worksheet protection flag to prevent accidental modification and to
# hide formulas if the locked and hidden format properties have been set.
#
sub protect {

    my $self = shift;

    $self->{_protect} = 1;

    # No password in XML format.
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


    # Check that cola are valid and store max and min values with default row.
    # NOTE: This isn't strictly correct. Excel only seems to set the dims
    #       for formatted/hidden columns. Should be conservative at least.
    return -2 if $self->_check_dimensions( 0, $data[0] );
    return -2 if $self->_check_dimensions( 0, $data[1] );

    push @{ $self->{_colinfo} }, [@data];

    # Store the col sizes for use when calculating image vertices taking
    # hidden columns into account. Also store the column formats.
    #
    my $width = $data[4] ? 0 : $data[2];    # Set width to zero if col is hidden
    $width ||= 0;                           # Ensure width isn't undef.
    my $format = $data[3];

    my ( $firstcol, $lastcol ) = @data;

    for my $col ( $firstcol .. $lastcol ) {
        $self->{_col_sizes}->{$col} = $width;
        $self->{_col_formats}->{$col} = $format if defined $format;
    }
}


###############################################################################
#
# set_selection()
#
# Set which cell or cells are selected in a worksheet: see also the
# sub _store_selection
#
sub set_selection {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    $self->{_selection} = [@_];
}


###############################################################################
#
# freeze_panes()
#
# Set panes and mark them as frozen. See also _store_panes().
#
sub freeze_panes {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    # Extra flag indicated a split and freeze.
    $self->{_frozen_no_split} = 0 if $_[4];

    $self->{_frozen} = 1;
    $self->{_panes}  = [@_];
}


###############################################################################
#
# split_panes()
#
# Set panes and mark them as split. See also _store_panes().
#
sub split_panes {

    my $self = shift;

    $self->{_frozen} = 0;
    $self->{_frozen_no_split}   = 0;
    $self->{_panes}  = [@_];
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

    $self->{_orientation} = 1;
}


###############################################################################
#
# set_landscape()
#
# Set the page orientation as landscape.
#
sub set_landscape {

    my $self = shift;

    $self->{_orientation} = 0;
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
# Set the colour of the worksheet colour.
#
sub set_tab_color {

    my $self  = shift;

    my $color = &Spreadsheet::WriteExcel::Format::_get_color($_[0]);
       $color = 0 if $color == 0x7FFF; # Default color.

    $self->{_tab_color} = $color;
}


###############################################################################
#
# set_paper()
#
# Set the paper type. Ex. 1 = US Letter, 9 = A4
#
sub set_paper {

    my $self = shift;

    $self->{_paper_size} = $_[0] || 0;
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

    $self->{_header} = $string;
    $self->{_margin_header} = $_[1] || 0.50;
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


    $self->{_footer} = $string;
    $self->{_margin_footer} = $_[1] || 0.50;
}


###############################################################################
#
# center_horizontally()
#
# Center the page horizontally.
#
sub center_horizontally {

    my $self = shift;

    if ( defined $_[0] ) {
        $self->{_hcenter} = $_[0];
    }
    else {
        $self->{_hcenter} = 1;
    }
}


###############################################################################
#
# center_vertically()
#
# Center the page horizontally.
#
sub center_vertically {

    my $self = shift;

    if ( defined $_[0] ) {
        $self->{_vcenter} = $_[0];
    }
    else {
        $self->{_vcenter} = 1;
    }
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

    my $self = shift;

    $self->{_margin_left} = defined $_[0] ? $_[0] : 0.75;
}


###############################################################################
#
# set_margin_right()
#
# Set the right margin in inches.
#
sub set_margin_right {

    my $self = shift;

    $self->{_margin_right} = defined $_[0] ? $_[0] : 0.75;
}


###############################################################################
#
# set_margin_top()
#
# Set the top margin in inches.
#
sub set_margin_top {

    my $self = shift;

    $self->{_margin_top} = defined $_[0] ? $_[0] : 1.00;
}


###############################################################################
#
# set_margin_bottom()
#
# Set the bottom margin in inches.
#
sub set_margin_bottom {

    my $self = shift;

    $self->{_margin_bottom} = defined $_[0] ? $_[0] : 1.00;
}


###############################################################################
#
# repeat_rows($first_row, $last_row)
#
# Set the rows to repeat at the top of each printed page. This is stored as
# <NamedRange> element.
#
sub repeat_rows {

    my $self = shift;

    my $row_min = $_[0];
    my $row_max = $_[1] || $_[0];    # Second row is optional

    my $area;

    # Convert the zero-indexed rows to R1:R2 notation.
    if ( $row_min == $row_max ) {
        $area = 'R' . ( $row_min + 1 );
    }
    else {
        $area = 'R' . ( $row_min + 1 ) . ':' . 'R' . ( $row_max + 1 );
    }

    # Build up the print area range "=Sheet2!R1:R2"
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

    my $area;

    # Convert the zero-indexed cols to C1:C2 notation.
    if ( $col_min == $col_max ) {
        $area = 'C' . ( $col_min + 1 );
    }
    else {
        $area = 'C' . ( $col_min + 1 ) . ':' . 'C' . ( $col_max + 1 );
    }

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

    $self->{_names}->{'Print_Area'} = $area;
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


    # Build up the print area range "=Sheet2!R1C1:R2C1"
    my $area = $self->_convert_name_area( $row1, $col1, $row2, $col2 );


    # Store the filter as a named range
    $self->{_names}->{'_FilterDatabase'} = $area;

    # Store the <Autofilter> information
    $area =~ s/[^!]+!//;    # Remove sheet name
    $self->{_autofilter} = $area;
    $self->{_filter_range} = [ $col1, $col2 ];
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


    # Check for a column reference in A1 notation and substitute.
    if ( $col =~ /^\D/ ) {
        my $col_letter = $col;
        ( undef, $col ) = xl_cell_to_rowcol( $col . '1' );

        croak "Invalid column '$col_letter'" if $col >= $self->{_xls_colmax};
    }


    my ( $col_first, $col_last ) = @{ $self->{_filter_range} };

    # Ignore column if it is outside filter range.
    return if $col < $col_first or $col > $col_last;


    my @tokens = split ' ', $expression;

    croak "Incorrect number of tokens in expression '$expression'"
      unless ( @tokens == 3 or @tokens == 7 );


    # We create an array slice to extract the operators from the arguments
    # and another to exclude the column placeholders.
    #
    # Index: 0 1 2  3  4 5 6
    #        x > 2
    #        x > 2 and x < 6

    my @slice1 = @tokens == 3 ? ( 1 ) : ( 1, 3, 5 );
    my @slice2 = @tokens == 3 ? ( 1, 2 ) : ( 1, 2, 3, 5, 6 );


    my %operators = (
        '==' => 'Equals',
        '='  => 'Equals',
        '=~' => 'Equals',
        'eq' => 'Equals',

        '!=' => 'DoesNotEqual',
        '!~' => 'DoesNotEqual',
        'ne' => 'DoesNotEqual',
        '<>' => 'DoesNotEqual',

        '>'  => 'GreaterThan',
        '>=' => 'GreaterThanOrEqual',
        '<'  => 'LessThan',
        '<=' => 'LessThanOrEqual',

        'and' => 'AutoFilterAnd',
        'or'  => 'AutoFilterOr',
        '&&'  => 'AutoFilterAnd',
        '||'  => 'AutoFilterOr',
    );


    for ( @tokens[@slice1] ) {
        if ( not exists $operators{$_} ) {
            croak "Unknown operator '$_'";
        }
    }


    for ( @tokens[@slice1] ) {
        for my $key ( keys %operators ) {
            s/^\Q$key\E$/$operators{$key}/i;
        }
    }

    $self->{_filter_cols}->{$col} = [ @tokens[@slice2] ];
    $self->{_filter_on} = 1;
}


###############################################################################
#
# _convert_name_area($first_row, $first_col, $last_row, $last_col)
#
# Convert zero indexed rows and columns to the R1C1 range required by worksheet
# named ranges, eg, "=Sheet2!R1C1:R2C1".
#
sub _convert_name_area {

    my $self = shift;

    my $row1 = $_[0];
    my $col1 = $_[1];
    my $row2 = $_[2];
    my $col2 = $_[3];

    my $range1 = '';
    my $range2 = '';
    my $area;


    # We need to handle some special cases that refer to rows or columns only.
    if ( $row1 == 0 and $row2 == $self->{_xls_rowmax} - 1 ) {
        $range1 = 'C' . ( $col1 + 1 );
        $range2 = 'C' . ( $col2 + 1 );
    }
    elsif ( $col1 == 0 and $col2 == $self->{_xls_colmax} - 1 ) {
        $range1 = 'R' . ( $row1 + 1 );
        $range2 = 'R' . ( $row2 + 1 );
    }
    else {
        $range1 = 'R' . ( $row1 + 1 ) . 'C' . ( $col1 + 1 );
        $range2 = 'R' . ( $row2 + 1 ) . 'C' . ( $col2 + 1 );
    }


    # A repeated range is only written once.
    if ( $range1 eq $range2 ) {
        $area = $range1;
    }
    else {
        $area = $range1 . ':' . $range2;
    }

    # Build up the print area range "=Sheet2!R1C1:R2C1"
    my $sheetname = $self->_quote_sheetname( $self->{_name} );
    $area = '=' . $sheetname . "!" . $area;


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

    my $self   = shift;
    my $option = $_[0];

    $option = 1 unless defined $option;    # Default to hiding printed gridlines

    if ( $option == 0 ) {
        $self->{_print_gridlines}  = 1;    # 1 = display, 0 = hide
        $self->{_screen_gridlines} = 1;
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
# print_gridlines()
#
# Turn on the printed gridlines.
#
sub print_gridlines {

    my $self = shift;

    $self->{_print_gridlines} = defined $_[0] ? $_[0] : 1;
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

    if ( defined $_[0] ) {
        $self->{_print_headers} = $_[0];
    }
    else {
        $self->{_print_headers} = 1;
    }
}


###############################################################################
#
# fit_to_pages($width, $height)
#
# Store the vertical and horizontal number of pages that will define the
# maximum area printed. See also _store_setup() and _store_wsbool() below.
#
sub fit_to_pages {

    my $self = shift;

    $self->{_fit_page}   = 1;
    $self->{_fit_width}  = $_[0] || 1;
    $self->{_fit_height} = $_[1] || 1;
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
# set_zoom($scale)
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

    # Turn off "fit to page" option
    $self->{_fit_page} = 0;

    $self->{_print_scale} = int $scale;
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
        croak "Not an array ref in call to write_row()$!";
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
# Write a comment to the specified row and column (zero indexed). The maximum
# comment size is 30831 chars. Excel5 probably accepts 32k-1 chars. However, it
# can only display 30831 chars. Excel 7 and 2000 will crash above 32k-1.
#
# In Excel 5 a comment is referred to as a NOTE.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long comment truncated to 30831 chars
#
sub write_comment {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }


    if ( @_ < 3 ) { return -1 }    # Check the number of args

    my $row     = $_[0];
    my $col     = $_[1];
    my $comment = $_[2];
    my $length  = length( $_[2] );
    my $error   = 0;
    my $max_len = 30831;             # Maintain same max as binary file.
    my $type    = 99;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    # String must be <= 30831 chars
    if ( $length > $max_len ) {
        $comment = substr( $comment, 0, $max_len );
        $error = -3;
    }


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );


    # Add a datatype to the cell if it doesn't already contain one.
    # This prevents an empty cell with a comment from being ignored.
    #
    if ( not $self->{_table}->[$row]->[$col] ) {
        $self->{_table}->[$row]->[$col] = [$type];
    }

    # Store the comment.
    $self->{_comment}->{$row}->{$col} = $comment;

    return $error;
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


    my $row  = $_[0];                              # Zero indexed row
    my $col  = $_[1];                              # Zero indexed column
    my $num  = $_[2];
    my $xf   = _XF( $self, $row, $col, $_[3] );    # The cell format
    my $type = 'n';                                # The data type

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    $self->{_table}->[$row]->[$col] = [ $type, $num, $xf ];

    return 0;
}


###############################################################################
#
# write_string ($row, $col, $string, $format, $html)
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

    my $row     = $_[0];                              # Zero indexed row
    my $col     = $_[1];                              # Zero indexed column
    my $str     = $_[2];
    my $xf      = _XF( $self, $row, $col, $_[3] );    # The cell format
    my $html    = $_[4] || 0;                         # Cell contains html text
    my $comment = '';                                 # Cell comment
    my $type    = 's';                                # The data type
    my $index;
    my $str_error = 0;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    if ( length $str > $self->{_xls_strmax} ) {    # LABEL must be < 32767 chars
        $str = substr( $str, 0, $self->{_xls_strmax} );
        $str_error = -3;
    }


    # TODO
    if ( not exists ${ $self->{_str_table} }->{$str} ) {
        ${ $self->{_str_table} }->{$str} = ${ $self->{_str_unique} }++;
    }


    ${ $self->{_str_total} }++;
    $index = ${ $self->{_str_table} }->{$str};


    $self->{_table}->[$row]->[$col] = [ $type, $index, $xf ];

    return $str_error;
}


###############################################################################
#
# write_html_string ($row, $col, $string, $format)
#
# Write a string to the specified row and column (zero indexed).
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 32767 chars
#
sub write_html_string {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    if ( @_ < 3 ) { return -1 }    # Check the number of args

    my $row  = $_[0];              # Zero indexed row
    my $col  = $_[1];              # Zero indexed column
    my $str  = $_[2];
    my $xf   = $_[3];              # The cell format
    my $html = 1;                  # Cell contains html text


    return $self->write_string( $row, $col, $str, $xf, $html );
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


    my $record = 0x0201;    # Record identifier
    my $length = 0x0006;    # Number of bytes to follow

    my $row  = $_[0];                              # Zero indexed row
    my $col  = $_[1];                              # Zero indexed column
    my $xf   = _XF( $self, $row, $col, $_[2] );    # The cell format
    my $type = 'b';                                # The data type

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

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
    my $value   = $_[4];           # The formula value.
    my $type    = 'f';             # The data type


    my $xf = _XF( $self, $row, $col, $_[3] );    # The cell format


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    # Remove the = sign if it exist.
    $formula =~ s/^=//;


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

    my $record = 0x0006;           # Record identifier
    my $length;                    # Bytes to follow

    my $row1    = $_[0];           # First row
    my $col1    = $_[1];           # First column
    my $row2    = $_[2];           # Last row
    my $col2    = $_[3];           # Last column
    my $formula = $_[4];           # The formula text string

    my $xf = _XF( $self, $row1, $col1, $_[5] );    # The cell format
    my $type = 99;                                 # The data type


    # Swap last row/col with first row/col as necessary
    ( $row1, $row2 ) = ( $row2, $row1 ) if $row1 > $row2;
    ( $col1, $col2 ) = ( $col1, $col2 ) if $col1 > $col2;


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row2, $col2 );


    # Define array range
    my $array_range;

    if ( $row1 == $row2 and $col1 == $col2 ) {
        $array_range = 'RC';
    }
    else {
        $array_range =
            xl_rowcol_to_cell( $row1, $col1 ) . ':'
          . xl_rowcol_to_cell( $row2, $col2 );
        $array_range = $self->_convert_formula( $row1, $col1, $array_range );
    }


    # Remove array formula braces and add = as required.
    $formula =~ s/^{(.*)}$/$1/;
    $formula =~ s/^([^=])/=$1/;


    # Convert A1 style references in the formula to R1C1 references
    $formula = $self->_convert_formula( $row1, $col1, $formula );

    $self->{_table}->[$row1]->[$col1] = [ $type, $formula, $xf, $array_range ];

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

    # Ensure this is a boolean vale for Window2
    $self->{_outline_on} = 1 if $self->{_outline_on};
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
    #
    my @args = @_;
    ( $args[3], $args[4] ) = ( $args[4], $args[3] ) if ref $args[3];


    my $row  = $args[0];                              # Zero indexed row
    my $col  = $args[1];                              # Zero indexed column
    my $url  = $args[2];                              # URL string
    my $str  = $args[3];                              # Alternative label
    my $xf   = _XF( $self, $row, $col, $args[4] );    # Tool tip
    my $tip  = $args[5];                              # XML data type
    my $type = 99;


    $url =~ s/^internal:/#/;    # Remove designators required by SWE.
    $url =~ s/^external://;     # Remove designators required by SWE.
    $str = $url unless defined $str;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    my $str_error = 0;


    $self->{_table}->[$row]->[$col] = [ $type, $url, $xf, $str, $tip ];

    return $str_error;
}


###############################################################################
#
# write_url_range($row1, $col1, $row2, $col2, $url, $string, $format)
#
# This is the more general form of write_url(). It allows a hyperlink to be
# written to a range of cells. This function also decides the type of hyperlink
# to be written. These are either, Web (http, ftp, mailto), Internal
# (Sheet1!A1) or external ('c:\temp\foo.xls#Sheet1!A1').
#
# See also write_url() above for a general description and return values.
#
sub write_url_range {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }

    # Check the number of args
    return -1 if @_ < 5;


    # Reverse the order of $string and $format if necessary. We work on a copy
    # in order to protect the callers args. We don't use "local @_" in case of
    # perl50005 threads.
    #
    my @args = @_;

    ( $args[5], $args[6] ) = ( $args[6], $args[5] ) if ref $args[5];

    my $url = $args[4];


    # Check for internal/external sheet links or default to web link
    return $self->_write_url_internal( @args ) if $url =~ m[^internal:];
    return $self->_write_url_external( @args ) if $url =~ m[^external:];
    return $self->_write_url_web( @args );
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

    my $row  = $_[0];                              # Zero indexed row
    my $col  = $_[1];                              # Zero indexed column
    my $str  = $_[2];
    my $xf   = _XF( $self, $row, $col, $_[3] );    # The cell format
    my $type = 'n';                                 # The data type


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    my $str_error = 0;
    my $date_time = $self->convert_date_time( $str );

    # If the date isn't valid then write it as a string.
    if ( not defined $date_time ) {
        $type      = 's';
        $str_error = -3;
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
# insert_bitmap($row, $col, $filename, $x, $y, $scale_x, $scale_y)
#
# Insert a 24bit bitmap image in a worksheet. The main record required is
# IMDATA but it must be proceeded by a OBJ record to define its position.
#
sub insert_bitmap {

    my $self = shift;


    # TODO Update for SpreadsheetML format

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
    my $height    = shift // 15;    # Row height.
    my $format    = shift;          # Format object.
    my $hidden    = shift // 0;     # Hidden flag.
    my $level     = shift // 0;     # Outline level.
    my $collapsed = shift // 0;     # Collapsed row.
    my $style;

    return unless defined $row;     # Ensure at least $row is specified.

    # Check that row and col are valid and store max and min values.
    return -2 if $self->_check_dimensions( $row, 0 );


    # If the height is 0 the row is hidden and the height is the default.
    if ( $height == 0 ) {
        $hidden = 1;
        $height = 15;
    }

    # Check for a format object.
    if ( ref $format ) {
        $style = $format->get_xf_index();
    }


    # Set the limits for the outline levels (0 <= x <= 7).
    $level = 0 if $level < 0;
    $level = 7 if $level > 7;

    if ( $level > $self->{_outline_row_level} ) {
        $self->{_outline_row_level} = $level;
    }


    # Store the row properties.
    $self->{_set_rows}->{$row} =
      [ $height, $style, $hidden, $level, $collapsed ];


    # Store the row sizes for use when calculating image vertices.
    # Also store the column formats.
    $self->{_row_sizes}->{$row} = $height;
    $self->{_row_formats}->{$row} = $format if defined $format;
}


###############################################################################
#
# merge_range($first_row, $first_col, $last_row, $last_col, $string, $format)
#
# This is a wrapper to ensure correct use of the merge_cells method, i.e. write
# the first cell of the range, write the formatted blank cells in the range and
# then call the merge_cells record. Failing to do the steps in this order will
# cause Excel 97 to crash.
#
sub merge_range {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $_[0] =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );
    }
    croak "Incorrect number of arguments" if @_ != 6;
    croak "Final argument must be a format object" unless ref $_[5];

    my $rwFirst  = $_[0];
    my $colFirst = $_[1];
    my $rwLast   = $_[2];
    my $colLast  = $_[3];
    my $string   = $_[4];
    my $format   = $_[5];


    # Excel doesn't allow a single cell to be merged
    croak "Can't merge single cell"
      if $rwFirst == $rwLast
          and $colFirst == $colLast;

    # Swap last row/col with first row/col as necessary
    ( $rwFirst,  $rwLast )  = ( $rwLast,  $rwFirst )  if $rwFirst > $rwLast;
    ( $colFirst, $colLast ) = ( $colLast, $colFirst ) if $colFirst > $colLast;


    # Check that column number is valid and store the max value
    return if $self->_check_dimensions( $rwLast, $colLast );


    # Store the merge range as a HoHoHoA
    $self->{_merge}->{$rwFirst}->{$colFirst} =
      [ $colLast - $colFirst, $rwLast - $rwFirst ];

    # Write the first cell
    return $self->write( $rwFirst, $colFirst, $string, $format );
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _XF()
#
# Returns an index to the XF record in the workbook.
#
# Note: this is a function, not a method.
#
sub _XF {

    # TODO $row and $col aren't actually required in the XML version and
    # should eventually be removed. They are required in the Biff version
    # to allow for row and col formats.

    my $self   = $_[0];
    my $row    = $_[1];
    my $col    = $_[2];
    my $format = $_[3];

    if ( ref( $format ) ) {
        return $format->get_xf_index();
    }
    else {
        return 0;    # 0x0F for Spreadsheet::WriteExcel
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
#
# This is an internal method that is used to filter elements of the array of
# pagebreaks used in the _store_hbreak() and _store_vbreak() methods. It:
#   1. Removes duplicate entries from the list.
#   2. Sorts the list.
#   3. Removes 0 from the list if present.
#
sub _sort_pagebreaks {

    my $self = shift;

    my %hash;
    my @array;

    @hash{@_} = undef;    # Hash slice to remove duplicates
    @array = sort { $a <=> $b } keys %hash;    # Numerical sort
    shift @array if $array[0] == 0;            # Remove zero

    # 1000 vertical pagebreaks appears to be an internal Excel 5 limit.
    # It is slightly higher in Excel 97/200, approx. 1026
    splice( @array, 1000 ) if ( @array > 1000 );

    return @array;
}


###############################################################################
#
# store_formula($formula)
#
# Pre-parse a formula. This is used in conjunction with repeat_formula()
# to repetitively rewrite a formula without re-parsing it.
#
sub store_formula {


    my $self = shift;

    # TODO Update for SpreadsheetML format
}


###############################################################################
#
# repeat_formula($row, $col, $formula, $format, ($pattern => $replacement,...))
#
# Write a formula to the specified row and column (zero indexed) by
# substituting $pattern $replacement pairs in the $formula created via
# store_formula(). This allows the user to repetitively rewrite a formula
# without the significant overhead of parsing.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub repeat_formula {

    my $self = shift;

    # TODO Update for SpreadsheetML format
}


###############################################################################
#
# _check_dimensions($row, $col, $ignore_row, $ignore_col)
#
# Check that $row and $col are valid and store max and min values for use in
# DIMENSIONS record. See, _store_dimensions().
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


    if ( not $ignore_row ) {

        if ( not defined $self->{_dim_rowmin} or $row < $self->{_dim_rowmin} ) {
            $self->{_dim_rowmin} = $row;
        }

        if ( not defined $self->{_dim_rowmax} or $row > $self->{_dim_rowmax} ) {
            $self->{_dim_rowmax} = $row;
        }
    }

    if ( not $ignore_col ) {

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
# _store_defcol()
#
# Write BIFF record DEFCOLWIDTH if COLINFO records are in use.
#
sub _store_defcol {

    my $self   = shift;
    my $record = 0x0055;    # Record identifier
    my $length = 0x0002;    # Number of bytes to follow

    my $colwidth = 0x0008;  # Default column width

    # TODO Update for SpreadsheetML format
}


###############################################################################
#
# _store_selection($first_row, $first_col, $last_row, $last_col)
#
# Write BIFF record SELECTION.
#
sub _store_selection {

    # TODO. Unused. Remove after refactoring.

    my $self   = shift;
    my $record = 0x001D;    # Record identifier
    my $length = 0x000F;    # Number of bytes to follow

    my $pnn     = $self->{_active_pane};    # Pane position
    my $rwAct   = $_[0];                    # Active row
    my $colAct  = $_[1];                    # Active column
    my $irefAct = 0;                        # Active cell ref
    my $cref    = 1;                        # Number of refs

    my $rwFirst  = $_[0];                   # First row in reference
    my $colFirst = $_[1];                   # First col in reference
    my $rwLast   = $_[2] || $rwFirst;       # Last  row in reference
    my $colLast  = $_[3] || $colFirst;      # Last  col in reference

    # Swap last row/col for first row/col as necessary
    if ( $rwFirst > $rwLast ) {
        ( $rwFirst, $rwLast ) = ( $rwLast, $rwFirst );
    }

    if ( $colFirst > $colLast ) {
        ( $colFirst, $colLast ) = ( $colLast, $colFirst );
    }


    # TODO Update for SpreadsheetML format
}


###############################################################################
#
# _store_externcount($count)
#
# Write BIFF record EXTERNCOUNT to indicate the number of external sheet
# references in a worksheet.
#
# Excel only stores references to external sheets that are used in formulas.
# For simplicity we store references to all the sheets in the workbook
# regardless of whether they are used or not. This reduces the overall
# complexity and eliminates the need for a two way dialogue between the formula
# parser the worksheet objects.
#
sub _store_externcount {

    # TODO. Unused. Remove after refactoring.

    my $self   = shift;
    my $record = 0x0016;    # Record identifier
    my $length = 0x0002;    # Number of bytes to follow

    my $cxals = $_[0];      # Number of external references

    # TODO Update for SpreadsheetML format
}


###############################################################################
#
# _store_externsheet($sheetname)
#
#
# Writes the Excel BIFF EXTERNSHEET record. These references are used by
# formulas. A formula references a sheet name via an index. Since we store a
# reference to all of the external worksheets the EXTERNSHEET index is the same
# as the worksheet index.
#
sub _store_externsheet {

    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    my $record = 0x0017;    # Record identifier
    my $length;             # Number of bytes to follow

    my $sheetname = $_[0];  # Worksheet name
    my $cch;                # Length of sheet name
    my $rgch;               # Filename encoding

    # References to the current sheet are encoded differently to references to
    # external sheets.
    #
    if ( $self->{_name} eq $sheetname ) {
        $sheetname = '';
        $length    = 0x02;    # The following 2 bytes
        $cch       = 1;       # The following byte
        $rgch      = 0x02;    # Self reference
    }
    else {
        $length = 0x02 + length( $_[0] );
        $cch    = length( $sheetname );
        $rgch = 0x03;         # Reference to a sheet in the current workbook
    }

    # TODO Update for SpreadsheetML format
}


###############################################################################
#
# _store_panes()
#
#
# Writes the Excel BIFF PANE record.
# The panes can either be frozen or thawed (unfrozen).
# Frozen panes are specified in terms of a integer number of rows and columns.
# Thawed panes are specified in terms of Excel's units for rows and columns.
#
sub _store_panes {

    # TODO. Unused. Remove after refactoring.

    my $self   = shift;
    my $record = 0x0041;    # Record identifier
    my $length = 0x000A;    # Number of bytes to follow

    my $y = $_[0] || 0;     # Vertical split position
    my $x = $_[1] || 0;     # Horizontal split position
    my $rwTop   = $_[2];    # Top row visible
    my $colLeft = $_[3];    # Leftmost column visible
    my $pnnAct  = $_[4];    # Active pane


    # Code specific to frozen or thawed panes.
    if ( $self->{_frozen} ) {

        # Set default values for $rwTop and $colLeft
        $rwTop   = $y unless defined $rwTop;
        $colLeft = $x unless defined $colLeft;
    }
    else {

        # Set default values for $rwTop and $colLeft
        $rwTop   = 0 unless defined $rwTop;
        $colLeft = 0 unless defined $colLeft;

        # Convert Excel's row and column units to the internal units.
        # The default row height is 12.75
        # The default column width is 8.43
        # The following slope and intersection values were interpolated.
        #
        $y = 20 * $y + 255;
        $x = 113.879 * $x + 390;
    }


    # Determine which pane should be active. There is also the undocumented
    # option to override this should it be necessary: may be removed later.
    #
    if ( not defined $pnnAct ) {
        $pnnAct = 0 if ( $x != 0 && $y != 0 );    # Bottom right
        $pnnAct = 1 if ( $x != 0 && $y == 0 );    # Top right
        $pnnAct = 2 if ( $x == 0 && $y != 0 );    # Bottom left
        $pnnAct = 3 if ( $x == 0 && $y == 0 );    # Top left
    }

    $self->{_active_pane} = $pnnAct;              # Used in _store_selection

    # TODO Update for SpreadsheetML format
}


###############################################################################
#
# _write_names()
#
# Write the <Worksheet> <Names> element.
#
sub _write_names {

    # TODO. Unused. Remove after refactoring.

    my $self = shift;


    if (    not keys %{ $self->{_names} }
        and not $self->{_repeat_rows}
        and not $self->{_repeat_cols} )
    {
        return;
    }


    if ( $self->{_repeat_rows} or $self->{_repeat_cols} ) {
        $self->{_names}->{Print_Titles} = '='
          . join ',',
          grep { /\S/ } $self->{_repeat_cols},
          $self->{_repeat_rows};
    }


    # Sort the <NamedRange> elements lexically and case insensitively.
    for my $key ( sort { lc $a cmp lc $b } keys %{ $self->{_names} } ) {

        my @attributes = (
            'NamedRange', 'ss:Name', $key, 'ss:RefersTo',
            $self->{_names}->{$key}
        );

        # Temp workaround to hide _FilterDatabase.
        # TODO. make this configurable later.
        if ( $key eq '_FilterDatabase' ) {
            push @attributes, 'ss:Hidden' => 1;
        }
    }
}


###############################################################################
#
# _store_pagebreaks()
#
# Store horizontal and vertical pagebreaks.
#
sub _store_pagebreaks {

    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    return
      if not @{ $self->{_hbreaks} }
          and not @{ $self->{_vbreaks} };
}


###############################################################################
#
# _store_protect()
#
# Set the Biff PROTECT record to indicate that the worksheet is protected.
#
sub _store_protect {

    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    # Exit unless sheet protection has been specified
    return unless $self->{_protect};

    my $record = 0x0012;    # Record identifier
    my $length = 0x0002;    # Bytes to follow

    my $fLock = $self->{_protect};    # Worksheet is protected

    # TODO Update for SpreadsheetML format
}


###############################################################################
#
# _size_col($col)
#
# Convert the width of a cell from user's units to pixels. Excel rounds the
# column width to the nearest pixel. Excel XML also scales the pixel value
# by 0.75.
#
sub _size_col {

    my $self  = shift;
    my $width = $_[0];

    # The relationship is different for user units less than 1.
    if ( $width < 1 ) {
        return 0.75 * int( $width * 12 );
    }
    else {
        return 0.75 * ( int( $width * 7 ) + 5 );
    }
}


###############################################################################
#
# _size_row($row)
#
# Convert the height of a cell from user's units to pixels. By interpolation
# the relationship is: y = 4/3x. Excel XML also scales the pixel value by 0.75.
#
sub _size_row {

    my $self   = shift;
    my $height = $_[0];

    return 0.75 * int( 4 / 3 * $height );
}


###############################################################################
#
# _store_zoom($zoom)
#
#
# Store the window zoom factor. This should be a reduced fraction but for
# simplicity we will store all fractions with a numerator of 100.
#
sub _store_zoom {

    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    # If scale is 100 we don't need to write a record
    return if $self->{_zoom} == 100;

    my $record = 0x00A0;    # Record identifier
    my $length = 0x0004;    # Bytes to follow

    # TODO Update for SpreadsheetML format
}


###############################################################################
#
# _store_comment
#
# Store the Excel 5 NOTE record. This format is not compatible with the Excel 7
# record.
#
sub _store_comment {

    # TODO. Unused. Remove after refactoring.

    my $self = shift;
    if ( @_ < 3 ) { return -1 }

    # TODO Update for SpreadsheetML format

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
# _write_autofilter()
#
# Write the <AutoFilter> element.
#
sub _write_autofilter {

    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    return unless $self->{_autofilter};

}


###############################################################################
#
# _write_autofilter_column()
#
# Write the <AutoFilterColumn> and <AutoFilterCondition> elements. The format
# of this is a little complicated.
#
sub _write_autofilter_column {

    # TODO. Unused. Remove after refactoring.

    my $self = shift;
    my @tokens;


    my ( $col_first, $col_last ) = @{ $self->{_filter_range} };

    my $prev_col = $col_first - 1;


    for my $col ( $col_first .. $col_last ) {

        # Check for rows with defined filter criteria.
        if ( defined $self->{_filter_cols}->{$col} ) {

            # Excel allows either one or two filter conditions

            # Single criterion.
            if ( @tokens == 2 ) {
                my ( $op, $value ) = @tokens;

            }

            # Double criteria, either 'And' or 'Or'.
            else {
                my ( $op1, $value1, $op2, $op3, $value3 ) = @tokens;

            }

        }
    }
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

    my $self                    = shift;
    my $published               = 0;
    my $conditional_calculation = 0;

    my @attributes = (
        'published'                         => $published,
        'enableFormatConditionsCalculation' => $conditional_calculation,
    );

    $self->{_writer}->emptyTag( 'sheetPr', @attributes );
}


###############################################################################
#
# _write_dimension()
#
# Write the <dimension> element. This specifies the range of cells in the
# worksheet. Ss a special case, empty spreadsheets use 'A1' as a range.
#
sub _write_dimension {

    my $self = shift;
    my $ref;

    if ( not defined $self->{_dim_rowmin} ) {

        # If the _dim_row_min is undefined then no dimensions have been set
        # and we use the default 'A1'.
        $ref = 'A1';
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
sub _write_sheet_view {

    my $self             = shift;
    my $tab_selected     = $self->{_selected};
    my $view             = 'pageLayout';
    my $workbook_view_id = 0;
    my @attributes       = ();

    if ( $tab_selected ) {
        push @attributes, ( 'tabSelected' => 1 );
    }

    push @attributes, ( 'workbookViewId' => $workbook_view_id );

    $self->{_writer}->emptyTag( 'sheetView', @attributes );

    # TODO. Add selection later.
    #$self->_write_selection();
    #$self->{_writer}->endTag( 'sheetView' );
}


###############################################################################
#
# _write_selection()
#
# Write the <selection> element.
#
sub _write_selection {

    my $self        = shift;
    my $active_cell = 'A1';
    my $sqref       = 'A1';

    my @attributes = (
        'activeCell' => $active_cell,
        'sqref'      => $sqref,
    );

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

    my @attributes = ( 'defaultRowHeight' => $default_row_height );

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
    my $min          = $_[0] // 0;    # First formatted column.
    my $max          = $_[1] // 0;    # Last formatted column.
    my $width        = $_[2];         # Col width in user units.
    my $format       = $_[3];         # Format object.
    my $hidden       = $_[4] // 0;    # Hidden flag.
    my $level        = $_[5] // 0;    # Outline level.
    my $collapsed    = $_[6] // 0;    # Outline level.
    my $style        = 0;
    my $custom_width = 1;

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


    # Check for a format object.
    if ( ref $format ) {
        $style = $format->get_xf_index();
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

    push @attributes, ( style       => 1 ) if $style;
    push @attributes, ( hidden      => 1 ) if $hidden;
    push @attributes, ( customWidth => 1 ) if $custom_width;


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
# _write_rows()
#
# Write out the worksheet data as a series of rows and cells.
#
sub _write_rows {

    my $self = shift;

    $self->_calculate_spans();

    for my $row_num ( $self->{_dim_rowmin} .. $self->{_dim_rowmax} ) {

        # Skip row if it doesn't contain row formatting or cell data.
        if ( !$self->{_set_rows}->{$row_num} && !$self->{_table}->[$row_num] ) {
            next;
        }

        # Write the cells if the row contains data.
        if ( my $row_ref = $self->{_table}->[$row_num] ) {
            my $span_index = int( $row_num / 16 );
            my $span       = $self->{_row_spans}->[$span_index];

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
        else {

            # Row attributes only.
            $self->_write_empty_row( $row_num, undef,
                @{ $self->{_set_rows}->{$row_num} } );
        }
    }
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
    my $height    = shift // 15;
    my $format    = shift;
    my $hidden    = shift // 0;
    my $level     = shift // 0;
    my $collapsed = shift // 0;
    my $empty_row = shift // 0;

    my @attributes = ( 'r' => $r + 1 );

    push @attributes, ( 'spans'        => $spans )  if defined $spans;
    push @attributes, ( 's'            => $format ) if $format;
    push @attributes, ( 'customFormat' => 1 )       if $format;
    push @attributes, ( 'ht'           => $height ) if $height != 15;
    push @attributes, ( 'hidden'       => 1 )       if $hidden;
    push @attributes, ( 'customHeight' => 1 )       if $height != 15;


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
    $self->_write_row( @_, 1 );
}


###############################################################################
#
# _write_cell()
#
# Write the <cell> element.
#
sub _write_cell {

    my $self  = shift;
    my $row   = shift;
    my $col   = shift;
    my $cell  = shift;
    my $type  = $cell->[0];
    my $value = $cell->[1];
    my $xf    = $cell->[2];


    my $range = xl_rowcol_to_cell( $row, $col );
    my @attributes = ( 'r' => $range );

    # Add the cell format index.
    if ( $xf ) {
        push @attributes, ( 's' => $xf );
    }


    if ( $type eq 'n' ) {
        $self->{_writer}->startTag( 'c', @attributes );
        $self->_write_cell_value( $value );
        $self->{_writer}->endTag( 'c' );
    }
    elsif ( $type eq 's' ) {
        push @attributes, ( 't' => 's' );

        $self->{_writer}->startTag( 'c', @attributes );
        $self->_write_cell_value( $value );
        $self->{_writer}->endTag( 'c' );
    }
    elsif ( $type eq 's' ) {
        push @attributes, ( 't' => 's' );

        $self->{_writer}->startTag( 'c', @attributes );
        $self->_write_cell_value( $value );
        $self->{_writer}->endTag( 'c' );
    }
    elsif ( $type eq 'f' ) {
        $self->{_writer}->startTag( 'c', @attributes );
        $self->_write_cell_formula( $value );
        $self->_write_cell_value( $cell->[3] );
        $self->{_writer}->endTag( 'c' );
    }
    elsif ( $type eq 'b' ) {
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
    my $value = shift // '';

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
    my $formula = shift // '';

    $self->{_writer}->dataElement( 'f', $formula );
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

    my $self   = shift;
    my $left   = 0.7;
    my $right  = 0.7;
    my $top    = 0.75;
    my $bottom = 0.75;
    my $header = 0.3;
    my $footer = 0.3;

    my @attributes = (
        'left'   => $left,
        'right'  => $right,
        'top'    => $top,
        'bottom' => $bottom,
        'header' => $header,
        'footer' => $footer,
    );

    $self->{_writer}->emptyTag( 'pageMargins', @attributes );
}

###############################################################################
#
# _write_page_setup()
#
# Write the <pageSetup> element.
#
sub _write_page_setup {

    my $self           = shift;
    my $paper_size     = 0;
    my $orientation    = 'portrait';
    my $horizontal_dpi = 4294967292;
    my $vertical_dpi   = 4294967292;

    my @attributes = (
        'paperSize'     => $paper_size,
        'orientation'   => $orientation,
        'horizontalDpi' => $horizontal_dpi,
        'verticalDpi'   => $vertical_dpi,
    );

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


1;


__END__


=head1 NAME

Worksheet - A writer class for Excel Worksheets.

=head1 SYNOPSIS

See the documentation for Excel::Writer::XLSX

=head1 DESCRIPTION

This module is used in conjunction with Excel::Writer::XLSX.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

� MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

