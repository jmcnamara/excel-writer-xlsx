package TestFunctions;

###############################################################################
#
# TestFunctions - Helper functions for Excel::XLSX::Writer test cases.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use 5.010000;
use Exporter;
use strict;
use warnings;
use Test::More;

our @ISA         = qw(Exporter);
our @EXPORT_OK   = ();
our @EXPORT      = qw(_expected_to_aref _got_to_aref _is_deep_diff );
our %EXPORT_TAGS = ();

our $VERSION = '0.01';


###############################################################################
#
# Read the __DATA__ section in the calling test program and return the data
# as an array ref with some data formatting.
#
sub _expected_to_aref {

    my @data;

    while ( <main::DATA> ) {
        next unless /\S/;
        chomp;
        s{/>$}{ />};
        s{^\s+}{};
        push @data, $_;
    }

    return \@data;
}


###############################################################################
#
# Convert an XML doc in a string to an array ref for test comparisons.
#
sub _got_to_aref {

    my $xml_str = shift;

    $xml_str =~ s/\n//;

    # Split the XML into chunks at element boundaries.
    my @data = split /(?<=>)(?=<)/, $xml_str;

    return \@data;
}


###############################################################################
#
# Use Test::Differences:: eq_or_diff() where available or else fall back to
# using Test::More::is_deeply().
#
sub _is_deep_diff {
    my ( $got, $expected, $caption, ) = @_;

    eval {
        require Test::Differences;
        Test::Differences->import();
    };

    if ( !$@ ) {
        eq_or_diff( $got, $expected, $caption, { context => 1 } );
    }
    else {
        is_deeply( $got, $expected, $caption );
    }

}


1;


__END__

