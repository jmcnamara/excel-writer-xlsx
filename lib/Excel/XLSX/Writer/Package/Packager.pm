package Excel::XLSX::Writer::Package::Packager;

###############################################################################
#
# Packager - A class for creating the Excel XLSX package.
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
use Excel::XLSX::Writer::Package::App;

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

    my $self = {
        _package_dir => '',
        _sheet_names  => [],
    };


    bless $self, $class;

    return $self;
}


###############################################################################
#
# _set_package_dir()
#
# Set the XLSX OPC package directory.
#
sub _set_package_dir {

    my $self = shift;

    $self->{_package_dir} = shift;
}


###############################################################################
#
# _set_sheet_names()
#
# Set the file names used in the XLSX package.
#
sub _set_sheet_names {

    my $self = shift;

    $self->{_sheet_names} = shift;
}


###############################################################################
#
# _create_package()
#
# Write the xml files that make up the XLXS OPC package.
#
sub _create_package {

    my $self = shift;

    $self->_write_app_file();

}


###############################################################################
#
# _write_app_file()
#
# Write the App.xml file.
#
sub _write_app_file {

    my $self = shift;
    my $dir  = $self->{_package_dir};
    my $app  = new Excel::XLSX::Writer::Package::App;

    mkdir $dir . '/docProps';

    for my $sheet_name (@{$self->{_sheet_names}}) {
        $app->_add_part_name($sheet_name);
    }

    $app->_set_xml_writer( $dir . '/docProps/App.xml' );
    $app->_assemble_xml_file();
}



1;


__END__

=pod

=head1 NAME

Packager - A class for creating the Excel XLSX package.

=head1 SYNOPSIS

See the documentation for L<Excel::XLSX::Writer>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::XLSX::Writer> to create an Excel XLSX container file.

From Wikipedia: I<The Open Packaging Conventions (OPC) is a container-file technology initially created by Microsoft to store a combination of XML and non-XML files that together form a single entity such as an Open XML Paper Specification (OpenXPS) document>. L<http://en.wikipedia.org/wiki/Open_Packaging_Conventions>.

At its simplest an Excel XLSX file contains the following elements:

     ____[Content_Types].xml
    |
    |____docProps
    | |____app.xml
    | |____core.xml
    |
    |____xl
    | |____workbook.xml
    | |____worksheets
    | | |____sheet1.xml
    | |
    | |____styles.xml
    | |
    | |____theme
    | | |____theme1.xml
    | |
    | |_____rels
    |   |____workbook.xml.rels
    |
    |_____rels
      |____.rels


The C<Excel::XLSX::Writer::Package::Packager> class co-ordinates the classes that represent the elements of the package and writes them into the XLSX file.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::XLSX::Writer>.

=cut
