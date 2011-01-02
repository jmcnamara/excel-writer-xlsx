package Excel::Writer::XLSX::Package::XMLwriterSimple;

###############################################################################
#
# XMLwriterSimple - TODO.
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
use Exporter;
use Carp;

our @ISA     = qw(Exporter);
our $VERSION = '0.02';


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;

    my $self = { _fh => shift };

    bless $self, $class;

    return $self;
}


###############################################################################
#
# xmlDecl()
#
# TODO
#
sub xmlDecl {

    my $self = shift;

    print { $self->{_fh} }
      qq(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n);

}


###############################################################################
#
# startTag()
#
# TODO
#
sub startTag {

    my $self       = shift;
    my $tag        = shift;
    my @attributes = @_;

    print { $self->{_fh} } "<$tag";

    while (@attributes) {
        my $key = shift @attributes;
        my $value = shift @attributes;

        print { $self->{_fh} } qq( $key="$value");
    }

    print { $self->{_fh} } ">";
}


###############################################################################
#
# endTag()
#
# TODO
#
sub endTag {

    my $self = shift;
    my $tag  = shift;

    print { $self->{_fh} } "</$tag>";
}


###############################################################################
#
# emptyTag()
#
# TODO
#
sub emptyTag {

    my $self       = shift;
    my $tag        = shift;
    my @attributes = @_;

    print { $self->{_fh} } "<$tag";

    while (@attributes) {
        my $key = shift @attributes;
        my $value = shift @attributes;

        print { $self->{_fh} } qq( $key="$value");
    }

    print { $self->{_fh} } " />";

}


###############################################################################
#
# dataElement()
#
# TODO
#
sub dataElement {

    my $self       = shift;
    my $tag        = shift;
    my $data       = shift;
    my @attributes = @_;

    print { $self->{_fh} } "<$tag";

    while (@attributes) {
        my $key = shift @attributes;
        my $value = shift @attributes;

        print { $self->{_fh} } qq( $key="$value");
    }

    for ($data) {
        s/&/&amp;/g;
        s/</&lt;/g;
        s/>/&gt;/g;
        s/"/&quot;/g;
    }

    print { $self->{_fh} } ">";
    print { $self->{_fh} } $data;
    print { $self->{_fh} } "</$tag>";
}


###############################################################################
#
# end()
#
# TODO
#
sub end {

    my $self = shift;

}


###############################################################################
#
# getOutput()
#
# TODO
#
sub getOutput {

    my $self = shift;

    return $self->{_fh};
}


1;


__END__

=pod

=head1 NAME

XMLwriterSimple - TODO.

=head1 SYNOPSIS

See the documentation for L<Excel::Writer::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Writer::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMX, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Writer::XLSX>.

=cut
