package Excel::Writer::XLSX::Package::RichValueTypes;

###############################################################################
#
# RichValueTypes - A class for writing the Excel XLSX rdRichValueTypes.xml file.
#
# Used in conjunction with Excel::Writer::XLSX
#
# Copyright 2000-2024, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.008002;
use strict;
use warnings;
use Carp;
use Excel::Writer::XLSX::Package::XMLwriter;

our @ISA     = qw(Excel::Writer::XLSX::Package::XMLwriter);
our $VERSION = '1.14';


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $fh    = shift;
    my $self  = Excel::Writer::XLSX::Package::XMLwriter->new( $fh );

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

    $self->xml_declaration;
    $self->_write_rv_types_info();


}


##############################################################################
#
# _write_rv_types_info()
#
# Write the <rvTypesInfo> element.
#
sub _write_rv_types_info {

    my $self = shift;
    my $xmlns =
      'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2';
    my $xmlns_mc =
      'http://schemas.openxmlformats.org/markup-compatibility/2006';
    my $mc_ignorable = 'x';
    my $xmlns_x = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    my @attributes = (
        'xmlns'        => $xmlns,
        'xmlns:mc'     => $xmlns_mc,
        'mc:Ignorable' => $mc_ignorable,
        'xmlns:x'      => $xmlns_x,
    );

    my $key_flags = [

        [ '_Self', [ 'ExcludeFromFile', 'ExcludeFromCalcComparison' ] ],
        [ '_DisplayString',          ['ExcludeFromCalcComparison'] ],
        [ '_Flags',                  ['ExcludeFromCalcComparison'] ],
        [ '_Format',                 ['ExcludeFromCalcComparison'] ],
        [ '_SubLabel',               ['ExcludeFromCalcComparison'] ],
        [ '_Attribution',            ['ExcludeFromCalcComparison'] ],
        [ '_Icon',                   ['ExcludeFromCalcComparison'] ],
        [ '_Display',                ['ExcludeFromCalcComparison'] ],
        [ '_CanonicalPropertyNames', ['ExcludeFromCalcComparison'] ],
        [ '_ClassificationId',       ['ExcludeFromCalcComparison'] ],

    ];

    $self->xml_start_tag( 'rvTypesInfo', @attributes );
    $self->xml_start_tag( 'global' );
    $self->xml_start_tag( 'keyFlags' );

    # Write the keyFlags element.
    for my $key_flag ( @$key_flags ) {
        my $key   = $key_flag->[0];
        my $flags = $key_flag->[1];

        $self->_write_key( $key, $flags );

    }

    $self->xml_end_tag( 'keyFlags' );
    $self->xml_end_tag( 'global' );
    $self->xml_end_tag( 'rvTypesInfo' );

}


##############################################################################
#
# _write_key()
#
# Write the <key> element.
#
sub _write_key {

    my $self  = shift;
    my $name  = shift;
    my $flags = shift;

    my @attributes = ( 'name' => $name, );

    $self->xml_start_tag( 'key', @attributes );

    for my $flag ( @$flags ) {

        # Write the flag element.
        $self->_write_flag( $flag );
    }


    $self->xml_end_tag( 'key' );
}


##############################################################################
#
# _write_flag()
#
# Write the <flag> element.
#
sub _write_flag {

    my $self  = shift;
    my $name  = shift;
    my $value = 1;

    my @attributes = (
        'name'  => $name,
        'value' => $value,
    );

    $self->xml_empty_tag( 'flag', @attributes );
}


1;


__END__

=pod

=head1 NAME

RichValueTypes - A class for writing the Excel XLSX rdRichValueTypes.xml file.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

(c) MM-MMXXIV, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<https://dev.perl.org/licenses/artistic.html> or the GNU General Public License v1.0 or later L<https://dev.perl.org/licenses/gpl1.html>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Writer::XLSX>.

=cut
