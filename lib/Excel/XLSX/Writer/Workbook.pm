package Excel::XLSX::Writer::Workbook;

###############################################################################
#
# Workbook - A writer class for Excel Workbooks.
#
#
# Used in conjunction with Excel::XLSX::Writer
#
# Copyright 2000-2010, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use Exporter;
use strict;
use warnings;
use Carp;
use FileHandle;
use Excel::XLSX::Writer::Worksheet;
use Excel::XLSX::Writer::Format;


use vars qw($VERSION @ISA);
@ISA = qw(Exporter);

$VERSION = '0.01';

###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $self;
    my $tmp_format = Excel::XLSX::Writer::Format->new();
    my $byte_order = $self->{_byte_order};


    $self->{_filename}          = $_[0] || '';
    $self->{_1904}              = 0;
    $self->{_activesheet}       = 0;
    $self->{_firstsheet}        = 0;
    $self->{_selected}          = 0;
    $self->{_xf_index}          = 21;            # 21 internal styles +1
    $self->{_fileclosed}        = 0;
    $self->{_biffsize}          = 0;
    $self->{_sheetname}         = "Sheet";
    $self->{_tmp_format}        = $tmp_format;
    $self->{_codepage}          = 0x04E4;
    $self->{_worksheets}        = [];
    $self->{_sheetnames}        = [];
    $self->{_formats}           = [];
    $self->{_palette}           = [];
    $self->{_lower_cell_limits} = 0;

    bless $self, $class;


    # Check for a filename unless it is an existing filehandle
    if ( not ref $self->{_filename} and $self->{_filename} eq '' ) {
        carp 'Filename required by Excel::XLSX::Writer->new()';
        return undef;
    }


    # If filename is a reference we assume that it is a valid filehandle.
    if ( ref $self->{_filename} ) {
        $self->{_filehandle} = $self->{_filename};
    }
    else {
        my $fh = FileHandle->new( '>' . $self->{_filename} );

        return undef unless defined $fh;

        # Set the output to utf8 in newer perls.
        if ( $] >= 5.008 ) {
            eval q(binmode $fh, ':utf8');
        }

        $self->{_filehandle} = $fh;
    }


    return $self;
}


###############################################################################
#
# close()
#
# Calls finalization methods.
#
sub close {

    my $self = shift;

    # In case close() is called twice, by user and by DESTROY.
    return if $self->{_fileclosed};

    # Test filehandle in case new() failed and the user didn't check.
    return unless defined $self->{_filehandle};

    $self->{_fileclosed} = 1;
    $self->_store_workbook();

    return close $self->{_filehandle};
}


###############################################################################
#
# DESTROY()
#
# Close the workbook if it hasn't already been explicitly closed.
#
sub DESTROY {

    my $self = shift;

    $self->close() if not $self->{_fileclosed};
}


###############################################################################
#
# sheets(slice,...)
#
# An accessor for the _worksheets[] array
#
# Returns: an optionally sliced list of the worksheet objects in a workbook.
#
sub sheets {

    my $self = shift;

    if ( @_ ) {

        # Return a slice of the array
        return @{ $self->{_worksheets} }[@_];
    }
    else {

        # Return the entire list
        return @{ $self->{_worksheets} };
    }
}


###############################################################################
#
# worksheets()
#
# An accessor for the _worksheets[] array.
# This method is now deprecated. Use the sheets() method instead.
#
# Returns: an array reference
#
sub worksheets {

    my $self = shift;

    return $self->{_worksheets};
}


###############################################################################
#
# add_worksheet($name)
#
# Add a new worksheet to the Excel workbook.
#
# Returns: reference to a worksheet object
#
sub add_worksheet {

    my $self = shift;
    my $name = $_[0] || "";

    # Check that sheetname is <= 31 chars (Excel limit).
    croak "Sheetname $name must be <= 31 chars" if length $name > 31;

    # Check that sheetname doesn't contain any invalid characters
    croak 'Invalid Excel character [:*?/\\] in worksheet name: ' . $name
      if $name =~ m{[:*?/\\]};

    my $index     = @{ $self->{_worksheets} };
    my $sheetname = $self->{_sheetname};

    if ( $name eq "" ) { $name = $sheetname . ( $index + 1 ) }

    # Check that the worksheet name doesn't already exist: a fatal Excel error.
    # The check must also exclude case insensitive matches.
    foreach my $tmp ( @{ $self->{_worksheets} } ) {
        if ( lc $name eq lc $tmp->get_name() ) {
            croak "Worksheet name '$name', with case ignored, "
              . "is already in use";
        }
    }


    # Porters take note, the following scheme of passing references to Workbook
    # data (in the \$self->{_foo} cases) instead of a reference to the Workbook
    # itself is a workaround to avoid circular references between Workbook and
    # Worksheet objects. Feel free to implement this in any way the suits your
    # language.
    #
    my @init_data = (
        $name,                  $index,
        $self->{_filehandle},   $self->{_indentation},
        \$self->{_activesheet}, \$self->{_firstsheet},
        $self->{_1904},         $self->{_lower_cell_limits},
    );

    my $worksheet = Excel::XLSX::Writer::Worksheet->new( @init_data );
    $self->{_worksheets}->[$index] = $worksheet;    # Store ref for iterator
    $self->{_sheetnames}->[$index] = $name;         # Store EXTERNSHEET names
    return $worksheet;
}


###############################################################################
#
# add_format(%properties)
#
# Add a new format to the Excel workbook. This adds an XF record and
# a FONT record. Also, pass any properties to the Format::new().
#
sub add_format {

    my $self = shift;

    my @init_data = ( $self->{_xf_index}, \$self->{_palette}, @_, );


    my $format = Excel::XLSX::Writer::Format->new( @init_data );

    $self->{_xf_index} += 1;
    push @{ $self->{_formats} }, $format;    # Store format reference

    return $format;
}


###############################################################################
#
# set_1904()
#
# Set the date system: 0 = 1900 (the default), 1 = 1904
#
sub set_1904 {

    my $self = shift;

    if ( defined( $_[0] ) ) {
        $self->{_1904} = $_[0];
    }
    else {
        $self->{_1904} = 1;
    }
}


###############################################################################
#
# get_1904()
#
# Return the date system: 0 = 1900, 1 = 1904
#
sub get_1904 {

    my $self = shift;

    return $self->{_1904};
}


###############################################################################
#
# set_custom_color()
#
# Change the RGB components of the elements in the colour palette.
#
sub set_custom_color {

    my $self = shift;


    # Match a HTML #xxyyzz style parameter
    if ( defined $_[1] and $_[1] =~ /^#(\w\w)(\w\w)(\w\w)/ ) {
        @_ = ( $_[0], hex $1, hex $2, hex $3 );
    }


    my $index = $_[0] || 0;
    my $red   = $_[1] || 0;
    my $green = $_[2] || 0;
    my $blue  = $_[3] || 0;

    my $aref = $self->{_palette};

    # Check that the colour index is the right range
    if ( $index < 8 or $index > 64 ) {
        carp "Color index $index outside range: 8 <= index <= 64";
        return 0;
    }

    # Check that the colour components are in the right range
    if (   ( $red < 0 or $red > 255 )
        || ( $green < 0 or $green > 255 )
        || ( $blue < 0  or $blue > 255 ) )
    {
        carp "Color component outside range: 0 <= color <= 255";
        return 0;
    }

    $index -= 8;    # Adjust colour index (wingless dragonfly)

    # Set the RGB value
    $aref->[$index] = [ $red, $green, $blue, 0 ];

    return $index + 8;
}


###############################################################################
#
# set_tempdir()
#
# Change the default temp directory used by _initialize() in Worksheet.pm.
#
sub set_tempdir {

    my $self = shift;

    # TODO Update for ExcelXML format
}


###############################################################################
#
# set_codepage()
#
# See also the _store_codepage method. This is used to store the code page, i.e.
# the character set used in the workbook.
#
sub set_codepage {

    my $self = shift;
    my $codepage = $_[0] || 1;
    $codepage = 0x04E4 if $codepage == 1;
    $codepage = 0x8000 if $codepage == 2;
    $self->{_codepage} = $codepage;
}


###############################################################################
#
# use_lower_cell_limits()
#
# TODO
#
sub use_lower_cell_limits {

    my $self = shift;

    croak "use_lower_cell_limits() must be called before add_worksheet()"
      if $self->sheets();

    $self->{_lower_cell_limits} = 1;
}


###############################################################################
#
# _store_workbook()
#
# Assemble worksheets into a workbook and send the BIFF data to an OLE
# storage.
#
sub _store_workbook {

    my $self = shift;

}


###############################################################################
#
# _store_all_xfs()
#
# Write all XF records.
#
sub _store_all_xfs {

    my $self = shift;

}


###############################################################################
#
# _store_externs()
#
# Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
# the NAME records.
#
sub _store_externs {

    my $self = shift;

}


###############################################################################
#
# _store_names()
#
# Write the NAME record to define the print area and the repeat rows and cols.
#
sub _store_names {

    my $self = shift;

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
# _write_workbook()
#
# Write <workbook> element.
#
sub _write_workbook {

    my $self   = shift;
    my $xmlns  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    my $xmlns_r =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    my @attributes = (
        'xmlns'   => $xmlns,
        'xmlns:r' => $xmlns_r,
    );

    $self->{_writer}->emptyTag( 'workbook', @attributes );
}

###############################################################################
#
# write_file_version()
#
# Write the <fileVersion> element.
#
sub _write_file_version {

    my $self          = shift;
    my $app_name      = 'xl';
    my $last_edited   = 4;
    my $lowest_edited = 4;
    my $rup_build     = 4505;

    my @attributes = (
        'appName'      => $app_name,
        'lastEdited'   => $last_edited,
        'lowestEdited' => $lowest_edited,
        'rupBuild'     => $rup_build,
    );

    $self->{_writer}->emptyTag( 'fileVersion', @attributes );
}

###############################################################################
#
# _write_workbook_pr()
#
# Write <workbookPr> element.
#
sub _write_workbook_pr {

    my $self                   = shift;
    my $date_1904              = 1;
    my $show_ink_annotation    = 0;
    my $auto_compress_pictures = 0;

    my @attributes = (
        'date1904'             => $date_1904,
        'showInkAnnotation'    => $show_ink_annotation,
        'autoCompressPictures' => $auto_compress_pictures,
    );

    $self->{_writer}->emptyTag( 'workbookPr', @attributes );
}

###############################################################################
#
# _write_book_views()
#
# Write <bookViews> element.
#
sub _write_book_views {

    my $self   = shift;

    $self->{_writer}->startTag( 'bookViews' );
    $self->_write_workbook_view();
    $self->{_writer}->endTag( 'bookViews' );
}

###############################################################################
#
# _write_workbook_view()
#
# Write <workbookView> element.
#
sub _write_workbook_view {

    my $self          = shift;
    my $x_window      = -20;
    my $y_window      = -20;
    my $window_width  = 34400;
    my $window_height = 20700;
    my $tab_ratio     = 500;

    my @attributes = (
        'xWindow'      => $x_window,
        'yWindow'      => $y_window,
        'windowWidth'  => $window_width,
        'windowHeight' => $window_height,
        'tabRatio'     => $tab_ratio,
    );

    $self->{_writer}->emptyTag( 'workbookView', @attributes );
}

###############################################################################
#
# _write_sheets()
#
# Write <sheets> element.
#
sub _write_sheets {

    my $self   = shift;

    $self->{_writer}->startTag( 'sheets' );
    $self->_write_sheet();
    $self->{_writer}->endTag( 'sheets' );
}


###############################################################################
#
# _write_sheet()
#
# Write <sheet> element.
#
sub _write_sheet {

    my $self     = shift;
    my $name     = 'Sheet1';
    my $sheet_id = 1;
    my $r_id     = 'rId1';

    my @attributes = (
        'name'    => $name,
        'sheetId' => $sheet_id,
        'r:id'    => $r_id,
    );

    $self->{_writer}->emptyTag( 'sheet', @attributes );
}

###############################################################################
#
# _write_calc_pr()
#
# Write <calcPr> element.
#
sub _write_calc_pr {

    my $self                 = shift;
    my $calc_id              = 130000;
    my $concurrent_calc      = 0;

    my @attributes = (
        'calcId'             => $calc_id,
        'concurrentCalc'     => $concurrent_calc,
    );

    $self->{_writer}->emptyTag( 'calcPr', @attributes );
}


###############################################################################
#
# _write_ext_lst()
#
# Write <extLst> element.
#
sub _write_ext_lst {

    my $self                 = shift;

    $self->{_writer}->startTag( 'extLst' );
    $self->_write_ext();
    $self->{_writer}->endTag( 'extLst' );
}


###############################################################################
#
# _write_ext()
#
# Write <ext> element.
#
sub _write_ext {

    my $self     = shift;
    my $xmlns_mx = 'http://schemas.microsoft.com/office/mac/excel/2008/main';
    my $uri      = 'http://schemas.microsoft.com/office/mac/excel/2008/main';

    my @attributes = (
        'xmlns:mx' => $xmlns_mx,
        'uri'      => $uri,
    );

    $self->{_writer}->startTag( 'ext', @attributes );
    $self->_write_mx_arch_id();
    $self->{_writer}->endTag( 'ext' );
}

###############################################################################
#
# _write_mx_arch_id()
#
# Write <mx:ArchID> element.
#
sub _write_mx_arch_id {

    my $self   = shift;
    my $Flags  = 2;

    my @attributes = ( 'Flags' => $Flags, );

    $self->{_writer}->emptyTag( 'mx:ArchID', @attributes );
}



1;


__END__


=head1 NAME

Workbook - A writer class for Excel Workbooks.

=head1 SYNOPSIS

See the documentation for Excel::XLSX::Writer

=head1 DESCRIPTION

This module is used in conjunction with Excel::XLSX::Writer.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
