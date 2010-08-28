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
use Carp;
use FileHandle;
use Excel::XLSX::Writer::XMLwriter;
use Excel::XLSX::Writer::Worksheet;
use Excel::XLSX::Writer::Format;


use vars qw($VERSION @ISA);
@ISA = qw(Excel::XLSX::Writer::XMLwriter Exporter);

$VERSION = '0.01';

###############################################################################
#
# new()
#
# Constructor. Creates a new Workbook object from a XMLwriter object.
#
sub new {

    my $class      = shift;
    my $self       = Excel::XLSX::Writer::XMLwriter->new();
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


    # Set colour palette.
    $self->set_palette_xl97();

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


    # Write the XML version.
    $self->_write_xml_directive( 0, 1, 0, 'xml', 'version', '1.0' );

    # Write the XML directive to make Windows open the file in Excel.
    $self->_write_xml_directive( 0, 1, 0, 'mso-application', 'progid',
        'Excel.Sheet' );

    # Write the XML namespaces.
    $self->_write_xml_start_tag(
        0,          1,
        1,          'Workbook',
        'xmlns:x',  'urn:schemas-microsoft-com:office:excel',
        'xmlns',    'urn:schemas-microsoft-com:office:spreadsheet',
        'xmlns:ss', 'urn:schemas-microsoft-com:office:spreadsheet',
    );


    $self->_store_all_xfs();


    # Ensure that at least one worksheet has been selected.
    if ( $self->{_activesheet} == 0 ) {
        @{ $self->{_worksheets} }[0]->{_selected} = 1;
    }

    # Calculate the number of selected worksheet tabs and call the finalization
    # methods for each worksheet
    foreach my $sheet ( @{ $self->{_worksheets} } ) {
        $self->{_selected}++ if $sheet->{_selected};
        $sheet->_close( $self->{_sheetnames} );
    }

    # Add Workbook globals
    $self->_store_codepage();
    $self->_store_externs();    # For print area and repeat rows
    $self->_store_names();      # For print area and repeat rows
    $self->_store_window1();
    $self->_store_1904();
    $self->_store_palette();


    # Close Workbook tag. WriteExcel _store_eof().
    $self->_write_xml_end_tag( 0, 1, 1, 'Workbook' );


    # Close the file
    #$self->{_filehandle}->close(); TODO
}


###############################################################################
#
# _store_all_xfs()
#
# Write all XF records.
#
sub _store_all_xfs {

    my $self = shift;
    my @attribs;

    $self->_write_xml_start_tag( 1, 1, 0, 'Styles' );

    # User defined XFs
    foreach my $format ( @{ $self->{_formats} } ) {

        $self->_write_xml_start_tag( 2, 1, 0, 'Style', 'ss:ID',
            's' . $format->get_xf_index() );


        # Write the <Alignment> properties if any
        if ( @attribs = $format->get_align_properties() ) {
            $self->_write_xml_element( 3, 1, 1, 'Alignment', @attribs );
        }


        # Write the <Borders> properties if any
        if ( @attribs = $format->get_border_properties() ) {
            $self->_write_xml_start_tag( 3, 1, 1, 'Borders' );

            for my $aref ( @attribs ) {
                $self->_write_xml_element( 4, 1, 0, 'Border', @$aref );
            }

            $self->_write_xml_end_tag( 3, 1, 1, 'Borders' );
        }


        # Write the <Font> properties if any
        if ( @attribs = $format->get_font_properties() ) {
            $self->_write_xml_element( 3, 1, 1, 'Font', @attribs );
        }


        # Write the <Interior> properties if any
        if ( @attribs = $format->get_interior_properties() ) {
            $self->_write_xml_element( 3, 1, 0, 'Interior', @attribs );
        }


        # Write the <NumberFormat> properties if any
        if ( @attribs = $format->get_num_format_properties() ) {
            $self->_write_xml_element( 3, 1, 0, 'NumberFormat', @attribs );
        }


        # Write the <Protection> properties if any
        if ( @attribs = $format->get_protection_properties() ) {
            $self->_write_xml_element( 3, 1, 0, 'Protection', @attribs );
        }


        $self->_write_xml_end_tag( 2, 1, 0, 'Style' );

    }

    $self->_write_xml_end_tag( 1, 1, 0, 'Styles' );

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

    # Create EXTERNCOUNT with number of worksheets
    $self->_store_externcount( scalar @{ $self->{_worksheets} } );

    # Create EXTERNSHEET for each worksheet
    foreach my $sheetname ( @{ $self->{_sheetnames} } ) {
        $self->_store_externsheet( $sheetname );
    }
}


###############################################################################
#
# _store_names()
#
# Write the NAME record to define the print area and the repeat rows and cols.
#
sub _store_names {

    my $self = shift;

    # Create the print area NAME records
    foreach my $worksheet ( @{ $self->{_worksheets} } ) {

        # Write a Name record if the print area has been defined
        if ( defined $worksheet->{_print_rowmin} ) {
            $self->_store_name_short(
                $worksheet->{_index},
                0x06,    # NAME type
                $worksheet->{_print_rowmin},
                $worksheet->{_print_rowmax},
                $worksheet->{_print_colmin},
                $worksheet->{_print_colmax}
            );
        }
    }


    # Create the print title NAME records
    foreach my $worksheet ( @{ $self->{_worksheets} } ) {

        my $rowmin = $worksheet->{_title_rowmin};
        my $rowmax = $worksheet->{_title_rowmax};
        my $colmin = $worksheet->{_title_colmin};
        my $colmax = $worksheet->{_title_colmax};

        # Determine if row + col, row, col or nothing has been defined
        # and write the appropriate record
        #
        if ( defined $rowmin && defined $colmin ) {

            # Row and column titles have been defined.
            # Row title has been defined.
            $self->_store_name_long(
                $worksheet->{_index},
                0x07,    # NAME type
                $rowmin,
                $rowmax,
                $colmin,
                $colmax
            );
        }
        elsif ( defined $rowmin ) {

            # Row title has been defined.
            $self->_store_name_short(
                $worksheet->{_index},
                0x07,    # NAME type
                $rowmin,
                $rowmax,
                0x00,
                0xff
            );
        }
        elsif ( defined $colmin ) {

            # Column title has been defined.
            $self->_store_name_short(
                $worksheet->{_index},
                0x07,    # NAME type
                0x0000,
                0x3fff,
                $colmin,
                $colmax
            );
        }
        else {

            # Print title hasn't been defined.
        }
    }
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

=head1 PATENT LICENSE

Software programs that read or write files that comply with the Microsoft specifications for the Office Schemas must include the following notice:

"This product may incorporate intellectual property owned by Microsoft Corporation. The terms and conditions upon which Microsoft is licensing such intellectual property may be found at http://msdn.microsoft.com/library/en-us/odcXMLRef/html/odcXMLRefLegalNotice.asp."

=head1 COPYRIGHT

© MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
