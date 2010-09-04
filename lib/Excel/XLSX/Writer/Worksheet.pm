package Excel::XLSX::Writer::Worksheet;

###############################################################################
#
# Worksheet - A writer class for Excel Worksheets.
#
#
# Used in conjunction with Excel::XLSX::Writer
#
# Copyright 2000-2010, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.010000;
use strict;
use warnings;
use Exporter;
use Carp;
use XML::Writer;
use Excel::XLSX::Writer::Format;
use Excel::XLSX::Writer::Utility qw(xl_cell_to_rowcol xl_rowcol_to_cell);

our @ISA     = qw(Exporter);
our $VERSION = '0.01';


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

    my $class = shift;
    my $self;
    my $rowmax = 1_048_576;
    my $colmax = 16_384;
    my $strmax = 32767;

    $self->{_name}              = $_[0];
    $self->{_index}             = $_[1];
    $self->{_filehandle}        = $_[2];
    $self->{_indentation}       = $_[3];
    $self->{_activesheet}       = $_[4];
    $self->{_firstsheet}        = $_[5];
    $self->{_1904}              = $_[6];
    $self->{_lower_cell_limits} = $_[7];

    $self->{_ext_sheets}  = [];
    $self->{_fileclosed}  = 0;
    $self->{_offset}      = 0;
    $self->{_xls_rowmax}  = $rowmax;
    $self->{_xls_colmax}  = $colmax;
    $self->{_xls_strmax}  = $strmax;
    $self->{_dim_rowmin}  = $rowmax + 1;
    $self->{_dim_rowmax}  = 0;
    $self->{_dim_colmin}  = $colmax + 1;
    $self->{_dim_colmax}  = 0;
    $self->{_dim_changed} = 0;
    $self->{_colinfo}     = [];
    $self->{_selection}   = [ 0, 0 ];
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
    $self->{_margin_head}   = 0.50;
    $self->{_margin_foot}   = 0.50;
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


    $self->{_datatypes} = {
        String   => 1,
        Number   => 2,
        DateTime => 3,
        Formula  => 4,
        Blank    => 5,
        HRef     => 6,
        Merge    => 7,
        Comment  => 8,
    };

    # Set older cell limits if required for backward compatibility.
    if ( $self->{_lower_cell_limits} ) {
        $self->{_xls_rowmax} = 65536;
        $self->{_xls_colmax} = 256;
    }


    bless $self, $class;
    $self->_initialize();
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

    $self->_write_xml_declaration;
    $self->_write_worksheet();
    # TODO.
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

    $self->_write_xml_start_tag( 1, 1, 0, 'Worksheet', 'ss:Name',
        $self->{_name} );

    # Write the Name elements such as print area and repeat rows.
    $self->_write_names();

    # Write the Table element and the child Row, Cell and Data elements.
    $self->_write_xml_table();

    # Write the worksheet page setup options.
    $self->_write_worksheet_options();

    # Store horizontal and vertical pagebreaks.
    $self->_store_pagebreaks();

    # Store autofilter information.
    $self->_write_autofilter();

    # Close Workbook tag. WriteExcel _store_eof().
    $self->_write_xml_end_tag( 1, 1, 1, 'Worksheet' );

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

    $self->{_selected} = 1;
    ${ $self->{_activesheet} } = $self->{_index};
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
# set_column($firstcol, $lastcol, $width, $format, $hidden, $autofit)
#
# Set the width of a single column or a range of columns.
# See also: _store_colinfo
#
sub set_column {

    my $self = shift;
    my $cell = $_[0];

    # Check for a cell reference in A1 notation and substitute row and column
    if ( $cell =~ /^\D/ ) {
        @_ = $self->_substitute_cellref( @_ );

        # Returned values $row1 and $row2 aren't required here. Remove them.
        shift @_;    # $row1
        splice @_, 1, 1;    # $row2
    }


    my ( $firstcol, $lastcol ) = @_;

    # Ensure at least $firstcol, $lastcol and $width
    return if @_ < 3;

    # Check that column number is valid and store the max value
    return if $self->_check_dimensions( 0, $lastcol );


    my $width   = $_[2];
    my $format  = _XF( $self, 0, 0, $_[3] );
    my $hidden  = $_[4];
    my $autofit = $_[5];

    if ( defined $width ) {
        $width = $self->_size_col( $_[2] );

        # The cell is hidden if the width is zero.
        $hidden = 1 if $width == 0;
    }


    foreach my $col ( $firstcol .. $lastcol ) {
        $self->{_set_cols}->{$col} = [ $width, $format, $hidden, $autofit ];
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

    $self->{_frozen} = 1;
    $self->{_panes}  = [@_];
}


###############################################################################
#
# thaw_panes()
#
# Set panes and mark them as unfrozen. See also _store_panes().
#
sub thaw_panes {

    my $self = shift;

    $self->{_frozen} = 0;
    $self->{_panes}  = [@_];
}


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
    $self->{_margin_head} = $_[1] || 0.50;
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
    $self->{_margin_foot} = $_[1] || 0.50;
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
# Center the page horinzontally.
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

    foreach my $token ( @$tokens ) {

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

    foreach my $token ( @$tokens ) {

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
    my $type = $self->{_datatypes}->{Comment};

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
    my $type = $self->{_datatypes}->{Number};      # The data type

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
    my $type    = $self->{_datatypes}->{String};      # The data type

    my $str_error = 0;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    if ( length $str > $self->{_xls_strmax} ) {    # LABEL must be < 32767 chars
        $str = substr( $str, 0, $self->{_xls_strmax} );
        $str_error = -3;
    }

    # Check if the cell already has a comment
    if ( $self->{_table}->[$row]->[$col] ) {
        $comment = $self->{_table}->[$row]->[$col]->[4];
    }


    $self->{_table}->[$row]->[$col] = [ $type, $str, $xf, $html, $comment ];

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
    my $type = $self->{_datatypes}->{Blank};       # The data type

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


    my $xf = _XF( $self, $row, $col, $_[3] );     # The cell format
    my $type = $self->{_datatypes}->{Formula};    # The data type


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );


    my $array_range = 'RC' if $formula =~ s/^{(.*)}$/$1/;

    # Add the = sign if it doesn't exist
    $formula =~ s/^([^=])/=$1/;


    # Convert A1 style references in the formula to R1C1 references
    $formula = $self->_convert_formula( $row, $col, $formula );


    $self->{_table}->[$row]->[$col] = [ $type, $formula, $xf, $array_range ];

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
    my $type = $self->{_datatypes}->{Formula};     # The data type


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
    my $type = $self->{_datatypes}->{HRef};


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
    my $type = $self->{_datatypes}->{DateTime};    # The data type


    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions( $row, $col );

    my $str_error = 0;
    my $date_time = $self->convert_date_time( $str );

    # If the date isn't valid then write it as a string.
    if ( not defined $date_time ) {
        $type      = $self->{_datatypes}->{String};
        $str_error = -3;
    }

    $self->{_table}->[$row]->[$col] = [ $type, $str, $xf ];

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

    # Can't store images in ExcelXML

    # TODO Update for ExcelXML format

}


###############################################################################
#
# set_row($row, $height, $XF, $hidden, $level)
#
# This method is used to set the height and XF format for a row.
#
sub set_row {

    my $self = shift;
    my $row  = $_[0];

    # Ensure at least $row and $height
    return if @_ < 2;

    # Check that row number is valid and store the max value
    return if $self->_check_dimensions( $row, 0 );


    my $height  = $_[1];
    my $format  = _XF( $self, 0, 0, $_[2] );
    my $hidden  = $_[3];
    my $autofit = $_[4];

    if ( $height ) {
        $height = $self->_size_row( $_[1] );

        # The cell is hidden if the width is zero.
        $hidden = 1 if $height == 0;
    }


    $self->{_set_rows}->{$row} = [ $height, $format, $hidden, $autofit ];
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

    # TODO Update for ExcelXML format
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

    # TODO Update for ExcelXML format
}
###############################################################################
#
# _check_dimensions($row, $col)
#
# Check that $row and $col are valid and store max and min values for use in
# DIMENSIONS record. See, _store_dimensions().
#
sub _check_dimensions {

    my $self = shift;
    my $row  = $_[0];
    my $col  = $_[1];

    if ( $row >= $self->{_xls_rowmax} ) { return -2 }
    if ( $col >= $self->{_xls_colmax} ) { return -2 }

    $self->{_dim_changed} = 1;

    if ( $row < $self->{_dim_rowmin} ) { $self->{_dim_rowmin} = $row }
    if ( $row > $self->{_dim_rowmax} ) { $self->{_dim_rowmax} = $row }
    if ( $col < $self->{_dim_colmin} ) { $self->{_dim_colmin} = $col }
    if ( $col > $self->{_dim_colmax} ) { $self->{_dim_colmax} = $col }

    return 0;
}


###############################################################################
#
# _store_window2()
#
# Write BIFF record Window2.
#
sub _store_window2 {

    use integer;    # Avoid << shift bug in Perl 5.6.0 on HP-UX

    my $self   = shift;
    my $record = 0x023E;    # Record identifier
    my $length = 0x000A;    # Number of bytes to follow

    my $grbit   = 0x00B6;        # Option flags
    my $rwTop   = 0x0000;        # Top row visible in window
    my $colLeft = 0x0000;        # Leftmost column visible in window
    my $rgbHdr  = 0x00000000;    # Row/column heading and gridline color

    # The options flags that comprise $grbit
    my $fDspFmla       = 0;                             # 0 - bit
    my $fDspGrid       = $self->{_screen_gridlines};    # 1
    my $fDspRwCol      = 1;                             # 2
    my $fFrozen        = $self->{_frozen};              # 3
    my $fDspZeros      = 1;                             # 4
    my $fDefaultHdr    = 1;                             # 5
    my $fArabic        = 0;                             # 6
    my $fDspGuts       = $self->{_outline_on};          # 7
    my $fFrozenNoSplit = 0;                             # 0 - bit
    my $fSelected      = $self->{_selected};            # 1
    my $fPaged         = 1;                             # 2

    # TODO Update for ExcelXML format
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

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_colinfo($firstcol, $lastcol, $width, $format, $autofit)
#
# Write XML <Column> elements to define column widths.
#
#
sub _store_colinfo {
    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    # Extract only the columns that have been defined.
    my @cols = sort { $a <=> $b } keys %{ $self->{_set_cols} };
    return unless @cols;

    my @attribs;
    my $previous = -1;
    my $span     = 0;

    for my $col ( @cols ) {
        if ( not $span ) {
            my $width   = $self->{_set_cols}->{$col}->[0];
            my $format  = $self->{_set_cols}->{$col}->[1];
            my $hidden  = $self->{_set_cols}->{$col}->[2];
            my $autofit = $self->{_set_cols}->{$col}->[3] || 0;

            push @attribs, "ss:Index",   $col + 1      if $col != $previous + 1;
            push @attribs, "ss:StyleID", "s" . $format if $format;
            push @attribs, "ss:Hidden",  $hidden       if $hidden;
            push @attribs, "ss:AutoFitWidth", $autofit;
            push @attribs, "ss:Width", $width if $width;

            # Note. "Overview of SpreadsheetML" states that the ss:Index
            # attribute is implicit in a Column element directly following a
            # Column element with an ss:Span attribute. However Excel doesn't
            # comply. In order to test directly against Excel we follow suit
            # and make ss:Index explicit. To get the implicit behaviour move
            # the next line outside the for() loop.
            $previous = $col;
        }

        # $previous = $col; # See note above.
        local $^W = 0;   # Ignore warnings about undefs in array ref comparison.

        # Check if the same attributes are shared over consecutive columns.
        if ( exists $self->{_set_cols}->{ $col + 1 }
            and join( "|", @{ $self->{_set_cols}->{$col} } ) eq
            join( "|", @{ $self->{_set_cols}->{ $col + 1 } } ) )
        {
            $span++;
            next;
        }

        push @attribs, "ss:Span", $span if $span;
        $self->_write_xml_element( 3, 1, 0, 'Column', @attribs );

        @attribs = ();
        $span    = 0;
    }
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


    # TODO Update for ExcelXML format
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

    # TODO Update for ExcelXML format
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

    # TODO Update for ExcelXML format
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

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_setup()
#
# Store the <WorksheetOptions> child element <PageSetup>.
#
sub _store_setup {
    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    # Write the <Layout> child element.
    my @layout;
    push @layout, 'x:Orientation', 'Landscape' if $self->{_orientation} == 0;
    push @layout, 'x:CenterHorizontal', 1 if $self->{_hcenter} == 1;
    push @layout, 'x:CenterVertical',   1 if $self->{_vcenter} == 1;
    push @layout, 'x:StartPageNumber', $self->{_start_page}
      if $self->{_page_start} > 0;


    # Write the <Header> child element.
    my @header;
    push @header, 'x:Margin', $self->{_margin_head}
      if $self->{_margin_head} != 0.5;
    push @header, 'x:Data', $self->{_header} if $self->{_header} ne '';


    # Write the <Footer> child element.
    my @footer;
    push @footer, 'x:Margin', $self->{_margin_foot}
      if $self->{_margin_foot} != 0.5;
    push @footer, 'x:Data', $self->{_footer} if $self->{_footer} ne '';


    # Write the <PageMargins> child element.
    my @margins;
    push @margins, 'x:Bottom', $self->{_margin_bottom}
      if $self->{_margin_bottom} != 1.00;
    push @margins, 'x:Left', $self->{_margin_left}
      if $self->{_margin_left} != 0.75;
    push @margins, 'x:Right', $self->{_margin_right}
      if $self->{_margin_right} != 0.75;
    push @margins, 'x:Top', $self->{_margin_top}
      if $self->{_margin_top} != 1.00;


    $self->_write_xml_element( 4, 1, 1, 'Layout',      @layout )  if @layout;
    $self->_write_xml_element( 4, 1, 1, 'Header',      @header )  if @header;
    $self->_write_xml_element( 4, 1, 1, 'Footer',      @footer )  if @footer;
    $self->_write_xml_element( 4, 1, 1, 'PageMargins', @margins ) if @margins;
}


###############################################################################
#
# _store_print()
#
# Store the <WorksheetOptions> child element <Print>.
#
sub _store_print {
    # TODO. Unused. Remove after refactoring.

    my $self = shift;


    if ( $self->{_fit_width} > 1 ) {
        $self->_write_xml_start_tag( 4, 0, 0, 'FitWidth' );
        $self->_write_xml_content( $self->{_fit_width} );
        $self->_write_xml_end_tag( 0, 1, 0, 'FitWidth' );
    }

    if ( $self->{_fit_height} > 1 ) {
        $self->_write_xml_start_tag( 4, 0, 0, 'FitHeight' );
        $self->_write_xml_content( $self->{_fit_height} );
        $self->_write_xml_end_tag( 0, 1, 0, 'FitHeight' );
    }


    # Print scale won't work without this.
    $self->_write_xml_element( 4, 1, 0, 'ValidPrinterInfo' );


    $self->_write_xml_element( 4, 1, 0, 'BlackAndWhite' )
      if $self->{_black_white};
    $self->_write_xml_element( 4, 1, 0, 'LeftToRight' ) if $self->{_page_order};
    $self->_write_xml_element( 4, 1, 0, 'DraftQuality' )
      if $self->{_draft_quality};


    if ( $self->{_paper_size} ) {
        $self->_write_xml_start_tag( 4, 0, 0, 'PaperSizeIndex' );
        $self->_write_xml_content( $self->{_paper_size} );
        $self->_write_xml_end_tag( 0, 1, 0, 'PaperSizeIndex' );
    }

    if ( $self->{_print_scale} != 100 ) {
        $self->_write_xml_start_tag( 4, 0, 0, 'Scale' );
        $self->_write_xml_content( $self->{_print_scale} );
        $self->_write_xml_end_tag( 0, 1, 0, 'Scale' );
    }


    $self->_write_xml_element( 4, 1, 0, 'Gridlines' )
      if $self->{_print_gridlines};
    $self->_write_xml_element( 4, 1, 0, 'RowColHeadings' )
      if $self->{_print_headers};
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


    $self->_write_xml_start_tag( 2, 1, 0, 'Names' );

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


        $self->_write_xml_element( 3, 1, 0, @attributes );

    }

    $self->_write_xml_end_tag( 2, 1, 0, 'Names' );

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

    $self->_write_xml_start_tag( 2, 1, 0, 'PageBreaks', 'xmlns',
        'urn:schemas-microsoft-com:' . 'office:excel' );


    if ( @{ $self->{_vbreaks} } ) {
        my @breaks = $self->_sort_pagebreaks( @{ $self->{_vbreaks} } );

        $self->_write_xml_start_tag( 3, 1, 0, 'ColBreaks' );

        for my $break ( @breaks ) {
            $self->_write_xml_start_tag( 4, 0, 0, 'ColBreak' );
            $self->_write_xml_start_tag( 0, 0, 0, 'Column' );
            $self->_write_xml_content( $break );
            $self->_write_xml_end_tag( 0, 0, 0, 'Column' );
            $self->_write_xml_end_tag( 0, 1, 0, 'ColBreak' );
        }

        $self->_write_xml_end_tag( 3, 1, 0, 'ColBreaks' );

    }

    if ( @{ $self->{_hbreaks} } ) {
        my @breaks = $self->_sort_pagebreaks( @{ $self->{_hbreaks} } );

        $self->_write_xml_start_tag( 3, 1, 0, 'RowBreaks' );

        for my $break ( @breaks ) {
            $self->_write_xml_start_tag( 4, 0, 0, 'RowBreak' );
            $self->_write_xml_start_tag( 0, 0, 0, 'Row' );
            $self->_write_xml_content( $break );
            $self->_write_xml_end_tag( 0, 0, 0, 'Row' );
            $self->_write_xml_end_tag( 0, 1, 0, 'RowBreak' );
        }

        $self->_write_xml_end_tag( 3, 1, 0, 'RowBreaks' );
    }

    $self->_write_xml_end_tag( 2, 1, 0, 'PageBreaks' );
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

    # TODO Update for ExcelXML format
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

    # TODO Update for ExcelXML format
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

    # TODO Update for ExcelXML format

}


###############################################################################
#
# New XML code
#
###############################################################################


###############################################################################
#
# _write_xml_table()
#
# Write the stored data into the <Table> element.
#
# TODO Add note about data structure
#
sub _write_xml_table {
    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    # Don't write <Table> element if it contains no data.
    return unless $self->{_dim_changed};


    $self->_write_xml_start_tag(
        2, 1, 0, 'Table', 'ss:ExpandedColumnCount', $self->{_dim_colmax} + 1,
        'ss:ExpandedRowCount', $self->{_dim_rowmax} + 1,
    );
    $self->_store_colinfo();

    # Write stored <Row> and <Cell> data
    $self->_write_xml_rows();

    $self->_write_xml_end_tag( 2, 1, 0, 'Table' );
}


###############################################################################
#
# _write_xml_rows()
#
# Write all <Row> elements.
#
sub _write_xml_rows {
    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    my @attribs;
    my $previous = -1;
    my $span     = 0;

    for my $row ( 0 .. $self->{_dim_rowmax} ) {

        next unless $self->{_set_rows}->{$row} or $self->{_table}->[$row];

        if ( not $span ) {
            my $height  = $self->{_set_rows}->{$row}->[0];
            my $format  = $self->{_set_rows}->{$row}->[1];
            my $hidden  = $self->{_set_rows}->{$row}->[2];
            my $autofit = $self->{_set_rows}->{$row}->[3] || 0;

            push @attribs, "ss:Index", $row + 1 if $row != $previous + 1;
            push @attribs, "ss:AutoFitHeight", $autofit if $height or $autofit;
            push @attribs, "ss:Height",        $height  if $height;
            push @attribs, "ss:Hidden",        $hidden  if $hidden;
            push @attribs, "ss:StyleID", "s" . $format if $format;

            # See ss:Index note in _store_colinfo
            $previous = $row;
        }

        # $previous = $row; # See ss:Index note in _store_colinfo
        local $^W = 0;   # Ignore warnings about undefs in array ref comparison.

        # Check if the same attributes are shared over consecutive columns.
        if (    not $self->{_table}->[$row]
            and not $self->{_table}->[ $row + 1 ]
            and exists $self->{_set_rows}->{$row}
            and exists $self->{_set_rows}->{ $row + 1 }
            and join( "|", @{ $self->{_set_rows}->{$row} } ) eq
            join( "|", @{ $self->{_set_rows}->{ $row + 1 } } ) )
        {
            $span++;
            next;
        }

        push @attribs, "ss:Span", $span if $span;

        # Write <Row> with <Cell> data or formatted <Row> without <Cell> data.
        #
        if ( my $row_ref = $self->{_table}->[$row] ) {
            $self->_write_xml_start_tag( 3, 1, 0, 'Row', @attribs );

            my $col = 0;
            $self->{prev_col} = -1;

            for my $col_ref ( @$row_ref ) {
                $self->_write_xml_cell( $row, $col ) if $col_ref;
                $col++;
            }
            $self->_write_xml_end_tag( 3, 1, 0, 'Row' );
        }
        else {
            $self->_write_xml_element( 3, 1, 0, 'Row', @attribs );
        }


        @attribs = ();
        $span    = 0;
    }
}


###############################################################################
#
# _write_xml_cell()
#
# Write a <Cell> element start tag.
#
sub _write_xml_cell {
    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    my $row = $_[0];
    my $col = $_[1];

    my $datatype = $self->{_table}->[$row]->[$col]->[0];
    my $data     = $self->{_table}->[$row]->[$col]->[1];
    my $format   = $self->{_table}->[$row]->[$col]->[2];

    my @attribs;
    my $comment = '';


    ###########################################################################
    #
    # Only add the cell index if it doesn't follow another cell.
    #
    push @attribs, "ss:Index", $col + 1 if $col != $self->{prev_col} + 1;


    ###########################################################################
    #
    # Check for merged cells.
    #
    if (    exists $self->{_merge}->{$row}
        and exists $self->{_merge}->{$row}->{$col} )
    {
        my ( $across, $down ) = @{ $self->{_merge}->{$row}->{$col} };

        push @attribs, "ss:MergeAcross", $across if $across;
        push @attribs, "ss:MergeDown",   $down   if $down;

        # Fill the merge range to ensure that it doesn't contain any data types.
        for my $m_row ( 0 .. $down ) {
            for my $m_col ( 0 .. $across ) {
                next if $m_row == 0 and $m_col == 0;
                $self->{_table}->[ $row + $m_row ]->[ $col + $m_col ] = undef;
            }
        }

        # Fill the last col so that $self->{prev_col} is incremented correctly.
        my $type = $self->{_datatypes}->{Merge};
        $self->{_table}->[$row]->[ $col + $across ] = [$type];
    }


    ###########################################################################
    #
    # Check for cell comments.
    #
    if (    exists $self->{_comment}->{$row}
        and exists $self->{_comment}->{$row}->{$col} )
    {
        $comment = $self->{_comment}->{$row}->{$col};
    }


    # Add the format attribute.
    push @attribs, "ss:StyleID", "s" . $format if $format;


    # Add to the attribute list for data types with additional options
    if ( $datatype == $self->{_datatypes}->{Formula} ) {
        my $array_range = $self->{_table}->[$row]->[$col]->[3];

        push @attribs, "ss:ArrayRange", $array_range if $array_range;
        push @attribs, "ss:Formula", $data;
    }

    if ( $datatype == $self->{_datatypes}->{HRef} ) {
        push @attribs, "ss:HRef", $data;

        my $tip = $self->{_table}->[$row]->[$col]->[4];
        push @attribs, "x:HRefScreenTip", $tip if defined $tip;
    }


    ###########################################################################
    #
    # Write the <Cell> data for various data types.
    #

    # Write the Number data element
    if ( $datatype == $self->{_datatypes}->{Number} ) {
        $self->_write_xml_start_tag( 4, 1, 0, 'Cell', @attribs );
        $self->_write_xml_cell_data( 'Number', $data );
        $self->_write_xml_cell_comment( $comment ) if $comment;
        $self->_write_xml_end_tag( 4, 1, 0, 'Cell' );
    }


    # Write the String data element
    elsif ( $datatype == $self->{_datatypes}->{String} ) {
        my $html = $self->{_table}->[$row]->[$col]->[3];

        $self->_write_xml_start_tag( 4, 1, 0, 'Cell', @attribs );

        if   ( $html ) { $self->_write_xml_html_string( $data ); }
        else           { $self->_write_xml_cell_data( 'String', $data ); }

        $self->_write_xml_cell_comment( $comment ) if $comment;
        $self->_write_xml_end_tag( 4, 1, 0, 'Cell' );
    }


    # Write the DateTime data element
    elsif ( $datatype == $self->{_datatypes}->{DateTime} ) {
        $self->_write_xml_start_tag( 4, 1, 0, 'Cell', @attribs );
        $self->_write_xml_cell_data( 'DateTime', $data );
        $self->_write_xml_cell_comment( $comment ) if $comment;
        $self->_write_xml_end_tag( 4, 1, 0, 'Cell' );
    }


    # Write an empty Data element for a formula data
    elsif ( $datatype == $self->{_datatypes}->{Formula} ) {
        if ( $comment ) {
            $self->_write_xml_start_tag( 4, 1, 0, 'Cell', @attribs );
            $self->_write_xml_cell_comment( $comment );
            $self->_write_xml_end_tag( 4, 1, 0, 'Cell' );
        }
        else {
            $self->_write_xml_element( 4, 1, 0, 'Cell', @attribs );
        }
    }


    # Write the HRef data element
    elsif ( $datatype == $self->{_datatypes}->{HRef} ) {

        $self->_write_xml_start_tag( 4, 1, 0, 'Cell', @attribs );

        my $data = $self->{_table}->[$row]->[$col]->[3];
        my $type;

        # Match DateTime string.
        if ( $self->convert_date_time( $data ) ) {
            $type = 'DateTime';
        }

        # Match integer with leading zero(s)
        elsif ( $self->{_leading_zeros} and $data =~ /^0\d+$/ ) {
            $type = 'String';
        }

        # Match number.
        elsif ( $data =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/ ) {
            $type = 'Number';
        }

        # Default to string.
        else {
            $type = 'String';
        }

        $self->_write_xml_cell_comment( $comment ) if $comment;
        $self->_write_xml_cell_data( $type, $data );
        $self->_write_xml_end_tag( 4, 1, 0, 'Cell' );
    }


    # Write an empty Data element for a blank cell
    elsif ( $datatype == $self->{_datatypes}->{Blank} ) {
        if ( $comment ) {
            $self->_write_xml_start_tag( 4, 1, 0, 'Cell', @attribs );
            $self->_write_xml_cell_comment( $comment );
            $self->_write_xml_end_tag( 4, 1, 0, 'Cell' );
        }
        else {
            $self->_write_xml_element( 4, 1, 0, 'Cell', @attribs );
        }
    }

    # Write an empty Data element for an empty cell with a comment;
    elsif ( $datatype == $self->{_datatypes}->{Comment} ) {
        if ( $comment ) {
            $self->_write_xml_start_tag( 4, 1, 0, 'Cell', @attribs );
            $self->_write_xml_cell_comment( $comment );
            $self->_write_xml_end_tag( 4, 1, 0, 'Cell' );
        }
        else {
            $self->_write_xml_element( 4, 1, 0, 'Cell', @attribs );
        }
    }

    # Ignore merge cells
    elsif ( $datatype == $self->{_datatypes}->{Merge} ) {

        # Do nothing.
    }


    $self->{prev_col} = $col;
    return;
}


###############################################################################
#
# _write_xml_cell_data()
#
# Write a generic Data element.
#
sub _write_xml_cell_data {
    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    my $datatype = $_[0];
    my $data     = $_[1];

    $self->_write_xml_start_tag( 5, 0, 0, 'Data', 'ss:Type', $datatype );

    if ( $datatype eq 'Number' ) {
        $self->_write_xml_unencoded_content( $data );
    }
    else { $self->_write_xml_content( $data ) }

    $self->_write_xml_end_tag( 0, 1, 0, 'Data' );
}


###############################################################################
#
# _write_xml_html_string()
#
# Write a string Data element with html text.
#
sub _write_xml_html_string {
    # TODO. Unused. Remove after refactoring.

    my $self = shift;
    my $data = $_[0];

    $self->_write_xml_start_tag( 5, 0, 0, 'ss:Data', 'ss:Type', 'String',
        'xmlns', 'http://www.w3.org/TR/REC-html40' );

    $self->_write_xml_unencoded_content( $data );

    $self->_write_xml_end_tag( 0, 1, 0, 'ss:Data' );
}


###############################################################################
#
# _write_xml_cell_comment()
#
# Write a cell Comment element.
#
sub _write_xml_cell_comment {
    # TODO. Unused. Remove after refactoring.

    my $self    = shift;
    my $comment = $_[0];

    $self->_write_xml_start_tag( 5, 1, 0, 'Comment' );

    $self->_write_xml_start_tag( 6, 0, 0, 'ss:Data', 'xmlns',
        'http://www.w3.org/TR/REC-html40' );

    $self->_write_xml_unencoded_content( $comment );

    $self->_write_xml_end_tag( 0, 1, 0, 'ss:Data' );

    $self->_write_xml_end_tag( 5, 1, 0, 'Comment' );

}


###############################################################################
#
# _write_worksheet_options()
#
# Write the <WorksheetOptions> element if the worksheet options have changed.
#
sub _write_worksheet_options {
    # TODO. Unused. Remove after refactoring.

    my $self = shift;

    my ( $options_changed, $print_changed, $setup_changed ) =
      $self->_options_changed();

    return unless $options_changed;

    $self->_write_xml_start_tag( 2, 1, 0, 'WorksheetOptions', 'xmlns',
        'urn:schemas-microsoft-com:' . 'office:excel' );


    if ( $setup_changed ) {
        $self->_write_xml_start_tag( 3, 1, 0, 'PageSetup' );
        $self->_store_setup();
        $self->_write_xml_end_tag( 3, 1, 0, 'PageSetup' );
    }


    $self->_write_xml_element( 3, 1, 0, 'FitToPage' ) if $self->{_fit_page};


    if ( $print_changed ) {
        $self->_write_xml_start_tag( 3, 1, 0, 'Print' );
        $self->_store_print();
        $self->_write_xml_end_tag( 3, 1, 0, 'Print' );
    }

    $self->_write_xml_element( 3, 1, 0, 'DoNotDisplayGridlines' )
      if $self->{_screen_gridlines} == 0;

    $self->_write_xml_element( 3, 1, 0, 'FilterOn' ) if $self->{_filter_on};

    $self->_write_xml_end_tag( 2, 1, 0, 'WorksheetOptions' );
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
        or $self->{_margin_head} != 0.50
        or $self->{_margin_foot} != 0.50
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

    $self->_write_xml_start_tag( 2, 1, 0, 'AutoFilter', 'x:Range',
        $self->{_autofilter}, 'xmlns',
        'urn:schemas-microsoft-com:' . 'office:excel' );


    $self->_write_autofilter_column();

    $self->_write_xml_end_tag( 2, 1, 0, 'AutoFilter' );
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

            my @attribs = ( 'AutoFilterColumn' );

            # The col indices are relative to the first column
            push @attribs, "x:Index", $col + 1 - $col_first
              if $col != $prev_col + 1;
            push @attribs, "x:Type", 'Custom';
            $prev_col = $col;

            $self->_write_xml_start_tag( 3, 1, 0, @attribs );

            @tokens = @{ $self->{_filter_cols}->{$col} };


            # Excel allows either one or two filter conditions

            # Single criterion.
            if ( @tokens == 2 ) {
                my ( $op, $value ) = @tokens;

                $self->_write_xml_element( 4, 1, 0, 'AutoFilterCondition',
                    'x:Operator', $op, 'x:Value', $value );
            }

            # Double criteria, either 'And' or 'Or'.
            else {
                my ( $op1, $value1, $op2, $op3, $value3 ) = @tokens;

                # <AutoFilterAnd> or <AutoFilterOr>
                $self->_write_xml_start_tag( 4, 1, 0, $op2 );

                $self->_write_xml_element( 5, 1, 0, 'AutoFilterCondition',
                    'x:Operator', $op1, 'x:Value', $value1 );

                $self->_write_xml_element( 5, 1, 0, 'AutoFilterCondition',
                    'x:Operator', $op3, 'x:Value', $value3 );

                $self->_write_xml_end_tag( 4, 1, 0, $op2 );

            }

            $self->_write_xml_end_tag( 3, 1, 0, 'AutoFilterColumn' );
        }
    }
}


###############################################################################
#
# _quote_sheetname()
#
# Sheetnames used in references should be quoted if they contain any spaces,
# special characters or if the look like something that isn't a sheet name.
# However, the rules are complex so for now we just quote anything that doesn't
# look like a simple sheet name.
#
sub _quote_sheetname {

    my $self      = shift;
    my $sheetname = $_[0];


    if ( $sheetname =~ /^Sheet\d+$/ ) {
        return $sheetname;
    }
    else {
        return "'" . $sheetname . "'";
    }
}


###############################################################################
#
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_xml_declaration()
#
# Write the XML declaration.
#
sub _write_xml_declaration {

    my $self       = shift;
    my $encoding   = 'UTF-8';
    my $standalone = 1;

    $self->{_writer}->xmlDecl( $encoding, $standalone );
}

###############################################################################
#
# _write_worksheet()
#
# Write the <worksheet> element.
#
sub _write_worksheet {

    my $self   = shift;
    my $xmlns  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    my $xmlns_r =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    my $xmlns_mc =
      'http://schemas.openxmlformats.org/markup-compatibility/2006';
    my $xmlns_mv               = 'urn:schemas-microsoft-com:mac:vml';
    my $mc_ignorable           = 'mv';
    my $mc_preserve_attributes = 'mv:*';

    my @attributes = (
        'xmlns'                 => $xmlns,
        'xmlns:r'               => $xmlns_r,
        'xmlns:mc'              => $xmlns_mc,
        'xmlns:mv'              => $xmlns_mv,
        'mc:Ignorable'          => $mc_ignorable,
        'mc:PreserveAttributes' => $mc_preserve_attributes,
    );

    $self->{_writer}->startTag( 'worksheet', @attributes );
    $self->{_writer}->endTag( 'worksheet' );
}


###############################################################################
#
# _write_sheet_pr()
#
# Write the <sheetPr> element.
#
sub _write_sheet_pr {

    my $self                                 = shift;
    my $published                            = 0;
    my $enable_format_conditions_calculation = 0;

    my @attributes = (
        'published' => $published,
        'enableFormatConditionsCalculation' =>
          $enable_format_conditions_calculation,
    );

    $self->{_writer}->emptyTag( 'sheetPr', @attributes );
}


###############################################################################
#
# _write_dimension()
#
# Write the <dimension> element.
#
sub _write_dimension {

    my $self   = shift;
    my $searef = 'A1:B2';

    my @attributes = ( 'searef' => $searef, );

    $self->{_writer}->emptyTag( 'dimension', @attributes );
}


###############################################################################
#
# _write_sheet_views()
#
# Write the <sheetViews> element.
#
sub _write_sheet_views {

    my $self   = shift;

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
    my $tab_selected     = 1;
    my $view             = 'pageLayout';
    my $workbook_view_id = 0;

    my @attributes = (
        'tabSelected'    => $tab_selected,
        'view'           => $view,
        'workbookViewId' => $workbook_view_id,
    );

    $self->{_writer}->startTag( 'sheetView', @attributes );
    $self->_write_selection();
    $self->{_writer}->endTag( 'sheetView' );
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
    my $default_row_height = 13;

    my @attributes = (
        'baseColWidth'     => $base_col_width,
        'defaultRowHeight' => $default_row_height,
    );

    $self->{_writer}->emptyTag( 'sheetFormatPr', @attributes );
}


###############################################################################
#
# _write_sheet_data()
#
# Write the <sheetData> element.
#
sub _write_sheet_data {

    my $self   = shift;

    $self->{_writer}->startTag( 'sheetData' );
    $self->{_writer}->endTag( 'sheetData' );
}


###############################################################################
#
# _write_row()
#
# Write the <row> element.
#
sub _write_row {

    my $self   = shift;
    my $writer = $self->{_writer};
    my $r      = 1;
    my $spans  = '1:3';

    my @attributes = (
        'r'     => $r,
        'spans' => $spans,
    );

    $self->{_writer}->startTag( 'row', @attributes );

    #$self->_write_foo();
    $self->{_writer}->endTag( 'row' );
}


###############################################################################
#
# _write_cell()
#
# Write the <cell> element.
#
sub _write_cell {

    my $self   = shift;
    my $value  = shift;
    my $range  = 'A1';

    my @attributes = ( 'r' => $range, );

    $self->{_writer}->startTag( 'c', @attributes );
    $self->_write_value( $value );
    $self->{_writer}->endTag( 'c' );
}


###############################################################################
#
# _write_value()
#
# Write the cell value <v> element.
#
sub _write_value {

    my $self   = shift;
    my $value  = shift;

    $self->{_writer}->dataElement( 'v', $value );
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
    my $left   = 0.75;
    my $right  = 0.75;
    my $top    = 1;
    my $bottom = 1;
    my $header = 0.5;
    my $footer = 0.5;

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

    my $self                 = shift;
    my $mode                 = 1;
    my $one_page             = 0;
    my $w_scale              = 0;

    my @attributes = (
        'Mode'               => $mode,
        'OnePage'            => $one_page,
        'WScale'             => $w_scale,
    );

    $self->{_writer}->emptyTag( 'mx:PLV', @attributes );
}



1;


__END__


=head1 NAME

Worksheet - A writer class for Excel Worksheets.

=head1 SYNOPSIS

See the documentation for Excel::XLSX::Writer

=head1 DESCRIPTION

This module is used in conjunction with Excel::XLSX::Writer.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

