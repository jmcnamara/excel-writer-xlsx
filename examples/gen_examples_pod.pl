#!/usr/bin/perl

###############################################################################
#
# Simple utility to convert the example programs listed in the README file into
# a Pod doc for easier access via CPAN.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;

my %images;
main();

###############################################################################
#
# main()
#
# Convert the example programs listed in the README file into a Pod doc for
# easier access via CPAN.
#
sub main {

    my @examples;
    my $examples_dir = $ARGV[0] || '.';

    # Get the version from the local Excel::Writer::XLSX.
    require "$examples_dir/../lib/Excel/Writer/XLSX.pm";
    my $version = Excel::Writer::XLSX->VERSION();

    # Read the filenames and descriptions from the examples README file.
    open my $readme, '<', $examples_dir . '/README'
      or die "Couldn't open $examples_dir/README file: $!\n";

    while ( my $line = <$readme> ) {
        if ( $line =~ /^\w+.pl\s/ ) {
            chomp $line;
            my ( $filename, $description ) = split " ", $line, 2;
            push @examples, [ $filename, $description ];
        }
    }

    die "Didn't find example programs in $examples_dir/README\n"
      unless @examples;

    read_images();

    print_header( $version );
    print_index( @examples );

    for my $example ( @examples ) {
        my $filename = $example->[0];
        print_example( $examples_dir, $filename, $version );
    }

    print_footer();
}

###############################################################################
#
# print_header()
#
# Print the header section of the Pod documentation.
#
sub print_header {

    my $version = shift;

    # I just don't like here docs.
    print "package Excel::Writer::XLSX::Examples;\n\n";

    print '#' x 79, "\n";
    print "#\n";
    print "# Examples - Excel::Writer::XLSX examples.\n";
    print "#\n";

    print "# A documentation only module showing the examples that are\n";
    print "# included in the Excel::Writer::XLSX distribution. This\n";
    print "# file was generated automatically via the gen_examples_pod.pl\n";
    print "# program that is also included in the examples directory.\n";
    print "#\n";

    print "# Copyright 2000-2020, John McNamara, jmcnamara\@cpan.org\n";
    print "#\n";
    print "# Documentation after __END__\n";
    print "#\n\n";

    print "use strict;\n";
    print "use warnings;\n";
    print "\n";
    print "our \$VERSION = '$version';\n\n";

    print "1;\n";
    print "\n";
    print "__END__\n\n";

    print "=pod\n\n";

    print "=encoding ISO8859-1\n\n";

    print "=head1 NAME\n\n";

    print "Examples - Excel::Writer::XLSX example programs.\n\n";

    print "=head1 DESCRIPTION\n\n";

    print "This is a documentation only module showing the examples that are\n";
    print "included in the L<Excel::Writer::XLSX> distribution.\n\n";
    print "This file was auto-generated via the gen_examples_pod.pl\n";
    print "program that is also included in the examples directory.\n";
    print "\n";

}

###############################################################################
#
# print_index()
#
# Print an index to the example programs with the short descriptions from the
# README file and a link to the appropriate section.
#
sub print_index {

    my @examples = @_;
    my $count    = scalar @examples;

    print "=head1 Example programs\n\n";

    print "The following is a list of the $count example programs that are ";
    print "included in the Excel::Writer::XLSX distribution.\n\n";

    print "=over\n\n";

    for my $example ( @examples ) {
        print "=item * L<Example: ";
        print $example->[0];
        print "> ";
        print $example->[1];
        print "\n\n";
    }

    print "=back\n\n";
}

###############################################################################
#
# print_example()
#
# Print each example program in its own =head1 section with a short description
# extracted from the first comment section of at the start and the code
# in a Pod verbatim section.
#
sub print_example {

    my $examples_dir = shift;
    my $example      = shift;
    my $version      = shift;
    my $verbatim     = '';
    my $in_header    = 0;

    open my $example_fh, '<', $examples_dir . '/' . $example;

    if ( !defined $example_fh ) {
        warn "Couldn't open $examples_dir/$example: $!\n";
        return undef;
    }

    print "=head2 Example: $example\n\n";

    while ( my $line = <$example_fh> ) {
        $line =~ s/\r//;
        $verbatim .= '    ' . $line;

        # Ignore the most common copyright line.
        next if $line =~ m/jmcnamara/;

        # Look for the first comment section but ignore the #!perl shebang line.
        if ( $in_header == 0 && $line !~ m/perl/ && $line =~ m/^#/ ) {
            $in_header = 1;
        }

        # In the first comment section.
        if ( $in_header == 1 ) {

            # Unset flag when leaving the first comment section.
            $in_header++ if $line !~ m/^#/;

            # Remove the comment char and the first leading space. This maintain
            # any embedded verbatim like sections.
            $line =~ s/^#+[ ]{0,1}//;

            print $line;
        }
    }

    print_image_html( $example );

    print $verbatim, "\n\n";

    print 'Download this example: L<http://cpansearch.perl.org/src/JMCNAMARA/';
    print "Excel-Writer-XLSX-$version/examples/$example>\n\n";
}

###############################################################################
#
# print_footer()
#
# Print the footer section of the Pod documentation
#
sub print_footer {

    print "=head1 AUTHOR\n\n";

    print "John McNamara jmcnamara\@cpan.org\n\n";

    print "Contributed examples contain the original author's name.\n\n";

    print "=head1 COPYRIGHT\n\n";

    print "Copyright MM-MMXVI, John McNamara.\n\n";

    print "All Rights Reserved. This module is free software. It may be used, ";
    print "redistributed and/or modified under the same terms as Perl itself.";
    print "\n\n";

    print "=cut\n";

}

###############################################################################
#
# read_images()
#
# Read the images associated with examples from the end of this file.
#
sub read_images {

    while ( <DATA> ) {
        next unless /\S/;
        next if /^#/;
        chomp;
        $images{$_} = 1;
    }
}

###############################################################################
#
# print_image_html()
#
# Print an embedded html image in the Pod doc if one exists for the example.
#
sub print_image_html {

    my $example = shift;
    my $image   = $example;

    $image =~ s/pl$/jpg/;

    return unless exists $images{$image};

    my $url    = 'http://jmcnamara.github.io/excel-writer-xlsx/images/examples';
    my $width  = 640;
    my $height = 420;

    print "=begin html\n\n";

    print '<p><center>';
    print qq{<img src="$url/$image" };
    print qq{width="$width" };
    print qq{height="$height" };
    print qq{alt="Output from $example" />};
    print qq{</center></p>\n\n};

    print "=end html\n\n";

    print "Source code for this example:\n\n";
}

__END__
# Image files used in the documentation.
a_simple.jpg
array_formula.jpg
autofilter.jpg
autofit.jpg
bigfile.jpg
chart_area.jpg
chart_bar.jpg
chart_column.jpg
chart_line.jpg
chart_pie.jpg
chart_doughnut.jpg
chart_radar.jpg
chart_scatter.jpg
chart_secondary_axis.jpg
chart_stock.jpg
chart_data_table.jpg
chart_data_tools.jpg
chess.jpg
colors.jpg
comments1.jpg
comments2.jpg
copyformat.jpg
conditional_format.jpg
data_validate.jpg
date_time.jpg
defined_name.jpg
demo.jpg
diag_border.jpg
filehandle.jpg
formats.jpg
formula_result.jpg
headers.jpg
hide_row_col.jpg
hide_sheet.jpg
hyperlink1.jpg
images.jpg
indent.jpg
macros.jpg
merge1.jpg
merge2.jpg
merge3.jpg
merge4.jpg
merge5.jpg
merge6.jpg
outline.jpg
outline_collapsed.jpg
panes.jpg
properties.jpg
protection.jpg
regions.jpg
repeat.jpg
rich_strings.jpg
right_to_left.jpg
row_wrap.jpg
sales.jpg
shape1.jpg
shape2.jpg
shape3.jpg
shape4.jpg
shape5.jpg
shape6.jpg
shape7.jpg
shape8.jpg
shape_all.jpg
sparklines1.jpg
sparklines2.jpg
stats.jpg
stats_ext.jpg
stocks.jpg
tab_colors.jpg
tables.jpg
textwrap.jpg
unicode_2022_jp.jpg
unicode_8859_11.jpg
unicode_8859_7.jpg
unicode_big5.jpg
unicode_cp1251.jpg
unicode_cp1256.jpg
unicode_cyrillic.jpg
unicode_koi8r.jpg
unicode_list.jpg
unicode_polish_utf8.jpg
unicode_shift_jis.jpg
unicode_utf16.jpg
unicode_utf16_japan.jpg
write_arrays.jpg
write_handler1.jpg
write_handler2.jpg
write_handler3.jpg
write_handler4.jpg
