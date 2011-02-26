package Excel::Writer::XLSX::Workbook;

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
use IO::File;
use File::Temp 'tempdir';
use Archive::Zip;
use Excel::Writer::XLSX::Worksheet;
use Excel::Writer::XLSX::Format;
use Excel::Writer::XLSX::Package::Packager;
use Excel::Writer::XLSX::Package::XMLwriter;
use Excel::Writer::XLSX::Utility qw(xl_cell_to_rowcol xl_rowcol_to_cell);

our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
our $VERSION = '0.14';


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
    my $self  = Excel::Writer::XLSX::Package::XMLwriter->new();

    $self->{_filename}         = $_[0] || '';
    $self->{_1904}             = 0;
    $self->{_activesheet}      = 0;
    $self->{_firstsheet}       = 0;
    $self->{_selected}         = 0;
    $self->{_xf_index}         = 0;
    $self->{_fileclosed}       = 0;
    $self->{_biffsize}         = 0;
    $self->{_sheetname}        = "Sheet";
    $self->{_codepage}         = 0x04E4;
    $self->{_worksheets}       = [];
    $self->{_sheetnames}       = [];
    $self->{_formats}          = [];
    $self->{_palette}          = [];
    $self->{_font_count}       = 0;
    $self->{_num_format_count} = 0;
    $self->{_defined_names}    = [];
    $self->{_named_ranges}     = [];
    $self->{_custom_colors}    = [];

    # Structures for the shared strings data.
    $self->{_str_total}  = 0;
    $self->{_str_unique} = 0;
    $self->{_str_table}  = {};
    $self->{_str_array}  = [];


    bless $self, $class;

    # Add the default cell format.
    $self->add_format();


    # Check for a filename unless it is an existing filehandle
    if ( not ref $self->{_filename} and $self->{_filename} eq '' ) {
        carp 'Filename required by Excel::Writer::XLSX->new()';
        return undef;
    }


    # If filename is a reference we assume that it is a valid filehandle.
    if ( ref $self->{_filename} ) {
        $self->{_filehandle} = $self->{_filename};
    }
    else {
        my $fh = IO::File->new( $self->{_filename}, 'w' );

        return undef unless defined $fh;

        # TODO check if the FH needs to be binmoded for Archive::Zip.
        #eval q(binmode $fh);

        $self->{_filehandle} = $fh;
    }


    # Set colour palette.
    $self->set_color_palette();

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

    # Write the root workbook element.
    $self->_write_workbook();

    # Write the XLSX file version.
    $self->_write_file_version();

    # Write the workbook properties.
    $self->_write_workbook_pr();

    # Write the workbook view properties.
    $self->_write_book_views();

    # Write the worksheet names and ids.
    $self->_write_sheets();

    # Write the workbook defined names.
    $self->_write_defined_names();

    # Write the workbook calculation properties.
    $self->_write_calc_pr();

    # Write the workbook extension storage.
    #$self->_write_ext_lst();

    # Close the workbook tag.
    $self->{_writer}->endTag( 'workbook' );

    # Close the XM writer object and filehandle.
    $self->{_writer}->end();
    $self->{_writer}->getOutput()->close();
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

    local ( $@, $!, $^E, $? );

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
    for my $tmp ( @{ $self->{_worksheets} } ) {
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
        $name,
        $index,

        \$self->{_activesheet},
        \$self->{_firstsheet},

        \$self->{_str_total},
        \$self->{_str_unique},
        \$self->{_str_table},

        $self->{_1904},
        $self->{_palette},
    );

    my $worksheet = Excel::Writer::XLSX::Worksheet->new( @init_data );
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

    my @init_data = ( $self->{_xf_index}, @_, );


    my $format = Excel::Writer::XLSX::Format->new( @init_data );

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

    # Set the RGB value.
    my @rgb = ( $red, $green, $blue );
    $aref->[$index] = [@rgb];

    # Store the custom colors for the style.xml file.
    push @{ $self->{_custom_colors} }, sprintf "FF%02X%02X%02X", @rgb;

    return $index + 8;
}


###############################################################################
#
# set_color_palette()
#
# Sets the colour palette to the Excel defaults.
#
sub set_color_palette {

    my $self = shift;

    $self->{_palette} = [
        [ 0x00, 0x00, 0x00, 0x00 ],    # 8
        [ 0xff, 0xff, 0xff, 0x00 ],    # 9
        [ 0xff, 0x00, 0x00, 0x00 ],    # 10
        [ 0x00, 0xff, 0x00, 0x00 ],    # 11
        [ 0x00, 0x00, 0xff, 0x00 ],    # 12
        [ 0xff, 0xff, 0x00, 0x00 ],    # 13
        [ 0xff, 0x00, 0xff, 0x00 ],    # 14
        [ 0x00, 0xff, 0xff, 0x00 ],    # 15
        [ 0x80, 0x00, 0x00, 0x00 ],    # 16
        [ 0x00, 0x80, 0x00, 0x00 ],    # 17
        [ 0x00, 0x00, 0x80, 0x00 ],    # 18
        [ 0x80, 0x80, 0x00, 0x00 ],    # 19
        [ 0x80, 0x00, 0x80, 0x00 ],    # 20
        [ 0x00, 0x80, 0x80, 0x00 ],    # 21
        [ 0xc0, 0xc0, 0xc0, 0x00 ],    # 22
        [ 0x80, 0x80, 0x80, 0x00 ],    # 23
        [ 0x99, 0x99, 0xff, 0x00 ],    # 24
        [ 0x99, 0x33, 0x66, 0x00 ],    # 25
        [ 0xff, 0xff, 0xcc, 0x00 ],    # 26
        [ 0xcc, 0xff, 0xff, 0x00 ],    # 27
        [ 0x66, 0x00, 0x66, 0x00 ],    # 28
        [ 0xff, 0x80, 0x80, 0x00 ],    # 29
        [ 0x00, 0x66, 0xcc, 0x00 ],    # 30
        [ 0xcc, 0xcc, 0xff, 0x00 ],    # 31
        [ 0x00, 0x00, 0x80, 0x00 ],    # 32
        [ 0xff, 0x00, 0xff, 0x00 ],    # 33
        [ 0xff, 0xff, 0x00, 0x00 ],    # 34
        [ 0x00, 0xff, 0xff, 0x00 ],    # 35
        [ 0x80, 0x00, 0x80, 0x00 ],    # 36
        [ 0x80, 0x00, 0x00, 0x00 ],    # 37
        [ 0x00, 0x80, 0x80, 0x00 ],    # 38
        [ 0x00, 0x00, 0xff, 0x00 ],    # 39
        [ 0x00, 0xcc, 0xff, 0x00 ],    # 40
        [ 0xcc, 0xff, 0xff, 0x00 ],    # 41
        [ 0xcc, 0xff, 0xcc, 0x00 ],    # 42
        [ 0xff, 0xff, 0x99, 0x00 ],    # 43
        [ 0x99, 0xcc, 0xff, 0x00 ],    # 44
        [ 0xff, 0x99, 0xcc, 0x00 ],    # 45
        [ 0xcc, 0x99, 0xff, 0x00 ],    # 46
        [ 0xff, 0xcc, 0x99, 0x00 ],    # 47
        [ 0x33, 0x66, 0xff, 0x00 ],    # 48
        [ 0x33, 0xcc, 0xcc, 0x00 ],    # 49
        [ 0x99, 0xcc, 0x00, 0x00 ],    # 50
        [ 0xff, 0xcc, 0x00, 0x00 ],    # 51
        [ 0xff, 0x99, 0x00, 0x00 ],    # 52
        [ 0xff, 0x66, 0x00, 0x00 ],    # 53
        [ 0x66, 0x66, 0x99, 0x00 ],    # 54
        [ 0x96, 0x96, 0x96, 0x00 ],    # 55
        [ 0x00, 0x33, 0x66, 0x00 ],    # 56
        [ 0x33, 0x99, 0x66, 0x00 ],    # 57
        [ 0x00, 0x33, 0x00, 0x00 ],    # 58
        [ 0x33, 0x33, 0x00, 0x00 ],    # 59
        [ 0x99, 0x33, 0x00, 0x00 ],    # 60
        [ 0x99, 0x33, 0x66, 0x00 ],    # 61
        [ 0x33, 0x33, 0x99, 0x00 ],    # 62
        [ 0x33, 0x33, 0x33, 0x00 ],    # 63
    ];

    return 0;
}


###############################################################################
#
# set_tempdir()
#
# Change the default temp directory used by _initialize() in Worksheet.pm.
#
sub set_tempdir {

    my $self = shift;

    # TODO Update for SpreadsheetML format
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
# _store_workbook()
#
# Assemble worksheets into a workbook and send the BIFF data to an OLE
# storage.
#
sub _store_workbook {

    my $self     = shift;
    my $dir      = tempdir( CLEANUP => 1 );
    my $packager = Excel::Writer::XLSX::Package::Packager->new();
    my $zip      = Archive::Zip->new();


    # Add a default worksheet if non have been added.
    $self->add_worksheet() if not @{ $self->{_worksheets} };

    # Ensure that at least one worksheet has been selected.
    if ( $self->{_activesheet} == 0 ) {
        @{ $self->{_worksheets} }[0]->{_selected} = 1;
        @{ $self->{_worksheets} }[0]->{_hidden}   = 0;
    }

    # Calculate the number of selected sheet tabs and set the active sheet.
    for my $sheet ( @{ $self->{_worksheets} } ) {
        $self->{_selected}++ if $sheet->{_selected};
        $sheet->{_active} = 1 if $sheet->{_index} == $self->{_activesheet};
    }

    # Convert the SST strings data structure.
    $self->_prepare_sst_string_data();

    # Set the font index for the format objects.
    $self->_prepare_fonts();

    # Set the number format index for the format objects.
    $self->_prepare_num_formats();

    # Set the border index for the format objects.
    $self->_prepare_borders();

    # Set the fill index for the format objects.
    $self->_prepare_fills();

    # Set the defined names for the worsheets such as Print Titles.
    $self->_prepare_defined_names();

    # Package the workbook.
    $packager->_add_workbook( $self );
    $packager->_set_package_dir( $dir );
    $packager->_create_package();

    # Free up the Packager object.
    $packager = undef;

    # Add the files to the zip archive.
    $zip->addTree( $dir, '', sub { -f } );

    if ( $zip->writeToFileHandle( $self->{_filehandle} ) != 0 ) {
        carp 'Error writing zip container for xlsx file.';
    }
}


###############################################################################
#
# _prepare_sst_string_data()
#
# Convert the SST string data from a hash to an array.
#
sub _prepare_sst_string_data {

    my $self = shift;

    my @strings;
    $#strings = $self->{_str_unique} - 1;    # Pre-extend array

    while ( my $key = each %{ $self->{_str_table} } ) {
        $strings[ $self->{_str_table}->{$key} ] = $key;
    }

    # The SST data could be very large, free some memory (maybe).
    $self->{_str_table} = undef;
    $self->{_str_array} = \@strings;

}


###############################################################################
#
# _prepare_fonts()
#
# Iterate through the XF Format objects and give them an index to non-default
# font elements.
#
sub _prepare_fonts {

    my $self = shift;

    my %fonts;
    my $index = 0;

    for my $format ( @{ $self->{_formats} } ) {
        my $key = $format->get_font_key();

        if ( exists $fonts{$key} ) {

            # Font has already been used.
            $format->{_font_index} = $fonts{$key};
            $format->{_has_font}   = 0;
        }
        else {

            # This is a new font.
            $fonts{$key}           = $index;
            $format->{_font_index} = $index;
            $format->{_has_font}   = 1;
            $index++;
        }
    }

    $self->{_font_count} = $index;
}


###############################################################################
#
# _prepare_num_formats()
#
# Iterate through the XF Format objects and give them an index to non-default
# number format elements.
#
# User defined records start from index 0xA4.
#
sub _prepare_num_formats {

    my $self = shift;

    my %num_formats;
    my $index            = 164;
    my $num_format_count = 0;

    for my $format ( @{ $self->{_formats} } ) {
        my $num_format = $format->{_num_format};

        # Check if $num_format is an index to a built-in number format.
        # Also check for a string of zeros, which is a valid number format
        # string but would evaluate to zero.
        #
        if ( $num_format =~ m/^\d+$/ && $num_format !~ m/^0+\d/ ) {

            # Index to a built-in number format.
            $format->{_num_format_index} = $num_format;
            next;
        }


        if ( exists( $num_formats{$num_format} ) ) {

            # Number format has already been used.
            $format->{_num_format_index} = $num_formats{$num_format};
        }
        else {

            # Add a new number format.
            $num_formats{$num_format} = $index;
            $format->{_num_format_index} = $index;
            $index++;
            $num_format_count++;
        }
    }

    $self->{_num_format_count} = $num_format_count;
}


###############################################################################
#
# _prepare_borders()
#
# Iterate through the XF Format objects and give them an index to non-default
# border elements.
#
sub _prepare_borders {

    my $self = shift;

    my %borders;
    my $index = 0;

    for my $format ( @{ $self->{_formats} } ) {
        my $key = $format->get_border_key();

        if ( exists $borders{$key} ) {

            # Border has already been used.
            $format->{_border_index} = $borders{$key};
            $format->{_has_border}   = 0;
        }
        else {

            # This is a new border.
            $borders{$key}           = $index;
            $format->{_border_index} = $index;
            $format->{_has_border}   = 1;
            $index++;
        }
    }

    $self->{_border_count} = $index;
}


###############################################################################
#
# _prepare_fills()
#
# Iterate through the XF Format objects and give them an index to non-default
# fill elements.
#
# The user defined fill properties start from 2 since there are 2 default
# fills: patternType="none" and patternType="gray125".
#
sub _prepare_fills {

    my $self = shift;

    my %fills;
    my $index = 2;    # Start from 2. See above.

    # Add the default fills.
    $fills{'0:0:0'}  = 0;
    $fills{'17:0:0'} = 1;

    for my $format ( @{ $self->{_formats} } ) {

        # The following logical statements jointly take care of special cases
        # in relation to cell colours and patterns:
        # 1. For a solid fill (_pattern == 1) Excel reverses the role of
        #    foreground and background colours, and
        # 2. If the user specifies a foreground or background colour without
        #    a pattern they probably wanted a solid fill, so we fill in the
        #    defaults.
        #
        if (   $format->{_pattern} <= 1
            && $format->{_bg_color} != 0
            && $format->{_fg_color} == 0 )
        {
            $format->{_fg_color} = $format->{_bg_color};
            $format->{_bg_color} = 0;
            $format->{_pattern}  = 1;
        }

        if (   $format->{_pattern} <= 1
            && $format->{_bg_color} == 0
            && $format->{_fg_color} != 0 )
        {
            $format->{_bg_color} = 0;
            $format->{_pattern}  = 1;
        }

        my $key = $format->get_fill_key();

        if ( exists $fills{$key} ) {

            # Fill has already been used.
            $format->{_fill_index} = $fills{$key};
            $format->{_has_fill}   = 0;
        }
        else {

            # This is a new fill.
            $fills{$key}           = $index;
            $format->{_fill_index} = $index;
            $format->{_has_fill}   = 1;
            $index++;
        }
    }

    $self->{_fill_count} = $index;
}

###############################################################################
#
# _prepare_defined_names()
#
# Iterate through the worksheets and store any defined names. Stores the
# defined name for the Workbook.xml and the named ranges for App.xml.
#
# TODO. Currently only supports Repeat rows/cols.
#
sub _prepare_defined_names {

    my $self = shift;

    for my $sheet ( @{ $self->{_worksheets} } ) {


        # Check for Print Area settings.
        if ( $sheet->{_autofilter} ) {

            my $range  = $sheet->{_autofilter};
            my $hidden = 1;

            # Store the defined names.
            push @{ $self->{_defined_names} },
              [ '_xlnm._FilterDatabase', $sheet->{_index}, $range, $hidden ];

        }

        # Check for Print Area settings.
        if ( $sheet->{_print_area} ) {

            my $range = $sheet->{_print_area};

            # Store the defined names.
            push @{ $self->{_defined_names} },
              [ '_xlnm.Print_Area', $sheet->{_index}, $range ];

            # Store the named ranges.
            my $sheetname   = $self->_quote_sheetname( $sheet->{_name} );
            my $print_title = $sheetname . '!Print_Area';

            push @{ $self->{_named_ranges} }, $print_title;
        }

        # Check for repeat rows/cols. aka, Print Titles.
        if ( $sheet->{_repeat_cols} || $sheet->{_repeat_rows} ) {
            my $range = '';

            if ( $sheet->{_repeat_cols} && $sheet->{_repeat_rows} ) {
                $range = $sheet->{_repeat_cols} . ',' . $sheet->{_repeat_rows};
            }
            else {
                $range = $sheet->{_repeat_cols} . $sheet->{_repeat_rows};
            }

            # Store the defined names.
            push @{ $self->{_defined_names} },
              [ '_xlnm.Print_Titles', $sheet->{_index}, $range ];

            # Store the named ranges.
            my $sheetname   = $self->_quote_sheetname( $sheet->{_name} );
            my $print_title = $sheetname . '!Print_Titles';

            push @{ $self->{_named_ranges} }, $print_title;
        }

    }
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
# XML writing methods.
#
###############################################################################


###############################################################################
#
# _write_workbook()
#
# Write <workbook> element.
#
sub _write_workbook {

    my $self    = shift;
    my $schema  = 'http://schemas.openxmlformats.org';
    my $xmlns   = $schema . '/spreadsheetml/2006/main';
    my $xmlns_r = $schema . '/officeDocument/2006/relationships';

    my @attributes = (
        'xmlns'   => $xmlns,
        'xmlns:r' => $xmlns_r,
    );

    $self->{_writer}->startTag( 'workbook', @attributes );
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
    my $date_1904              = $self->{_1904};
    my $show_ink_annotation    = 0;
    my $auto_compress_pictures = 0;
    my $default_theme_version  = 124226;
    my @attributes;

    push @attributes, ( 'date1904' => 1 ) if $date_1904;
    push @attributes, ( 'defaultThemeVersion' => $default_theme_version );

    $self->{_writer}->emptyTag( 'workbookPr', @attributes );
}


###############################################################################
#
# _write_book_views()
#
# Write <bookViews> element.
#
sub _write_book_views {

    my $self = shift;

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
    my $x_window      = 240;
    my $y_window      = 15;
    my $window_width  = 16095;
    my $window_height = 9660;
    my $tab_ratio     = 500;
    my $active_tab    = $self->{_activesheet};
    my $first_sheet   = $self->{_firstsheet};

    my @attributes = (
        'xWindow'      => $x_window,
        'yWindow'      => $y_window,
        'windowWidth'  => $window_width,
        'windowHeight' => $window_height,
    );

    # Store the firstSheet attribute when it isn't the default.
    push @attributes, ( firstSheet => $first_sheet ) if $first_sheet > 0;

    # Store the activeTab attribute when it isn't the first sheet.
    push @attributes, ( activeTab => $active_tab ) if $active_tab > 0;

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
    my $id_num = 1;

    $self->{_writer}->startTag( 'sheets' );

    for my $worksheet ( @{ $self->{_worksheets} } ) {
        $self->_write_sheet( $worksheet->{_name}, $id_num++,
            $worksheet->{_hidden} );
    }

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
    my $name     = shift;
    my $sheet_id = shift;
    my $hidden   = shift;
    my $r_id     = 'rId' . $sheet_id;

    my @attributes = (
        'name'    => $name,
        'sheetId' => $sheet_id,
    );

    push @attributes, ( 'state' => 'hidden' ) if $hidden;
    push @attributes, ( 'r:id' => $r_id );


    $self->{_writer}->emptyTag( 'sheet', @attributes );
}


###############################################################################
#
# _write_calc_pr()
#
# Write <calcPr> element.
#
sub _write_calc_pr {

    my $self            = shift;
    my $calc_id         = 124519;
    my $concurrent_calc = 0;

    my @attributes = ( 'calcId' => $calc_id, );

    $self->{_writer}->emptyTag( 'calcPr', @attributes );
}


###############################################################################
#
# _write_ext_lst()
#
# Write <extLst> element.
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

    my $self  = shift;
    my $Flags = 2;

    my @attributes = ( 'Flags' => $Flags, );

    $self->{_writer}->emptyTag( 'mx:ArchID', @attributes );
}


##############################################################################
#
# _write_defined_names()
#
# Write the <definedNames> element.
#
sub _write_defined_names {

    my $self = shift;

    return unless @{ $self->{_defined_names} };

    $self->{_writer}->startTag( 'definedNames' );

    for my $aref ( @{ $self->{_defined_names} } ) {
        $self->_write_defined_name( $aref );
    }

    $self->{_writer}->endTag( 'definedNames' );
}


##############################################################################
#
# _write_defined_name()
#
# Write the <definedName> element.
#
sub _write_defined_name {

    my $self = shift;
    my $data = shift;

    my $name           = $data->[0];
    my $local_sheet_id = $data->[1];
    my $range          = $data->[2];
    my $hidden         = $data->[3];

    my @attributes = (
        'name'         => $name,
        'localSheetId' => $local_sheet_id,
    );

    push @attributes, ( 'hidden' => 1) if $hidden;

    $self->{_writer}->dataElement( 'definedName', $range, @attributes );
}


1;


__END__


=head1 NAME

Workbook - A writer class for Excel Workbooks.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMXI, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
