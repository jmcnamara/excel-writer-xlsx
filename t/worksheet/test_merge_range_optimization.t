###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2025, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use strict;
use warnings;

use Excel::Writer::XLSX;
use File::Temp qw(tempdir);

use Test::More tests => 3;


###############################################################################
#
# Tests setup.
#
# set_optimization() is a memory optimization and must not change the output,
# so the merges recorded with it on have to match the merges recorded with it
# off for the same input.
#
my $dir     = tempdir( CLEANUP => 1 );
my $caption = " \tWorksheet: merge_range() under set_optimization()";


###############################################################################
#
# Write a merge that spans several rows, cell by cell in row order, the way a
# streaming writer has to under set_optimization(): the first cell carries the
# value, the rest are padded with write_blank(), and merge_range() can only be
# called once the last cell of the range is reached.
#
sub _merges_for {

    my ( $filename, $optimize ) = @_;

    my $workbook = Excel::Writer::XLSX->new( $filename );
    $workbook->set_optimization() if $optimize;

    my $worksheet = $workbook->add_worksheet();
    my $format    = $workbook->add_format( align => 'center' );
    my $return;

    for my $row ( 0 .. 3 ) {

        # Column A holds a merge spanning all four rows.
        if ( $row == 0 ) {
            $worksheet->write( $row, 0, 'Merged', $format );
        }
        else {
            $worksheet->write_blank( $row, 0, $format );
            $return = $worksheet->merge_range( 0, 0, 3, 0, 'Merged', $format )
              if $row == 3;
        }

        # Columns B and C hold ordinary streamed data.
        $worksheet->write( $row, 1, "B$row" );
        $worksheet->write( $row, 2, "C$row" );
    }

    $workbook->close();

    my $xml = do {
        local $/;
        open my $fh, '-|', 'unzip', '-p', $filename, 'xl/worksheets/sheet1.xml'
          or die "Failed to unzip $filename: $!";
        <$fh>;
    };

    my @merges = $xml =~ /mergeCell ref="([^"]+)"/g;

    return ( $return, \@merges );
}


my ( $plain_return, $plain_merges ) =
  _merges_for( "$dir/plain.xlsx", 0 );

my ( $optimized_return, $optimized_merges ) =
  _merges_for( "$dir/optimized.xlsx", 1 );


###############################################################################
#
# Test that the merge is recorded in both modes, and identically.
#
is_deeply( $plain_merges, ['A1:A4'], $caption );

is_deeply( $optimized_merges, $plain_merges, $caption );

is( $optimized_return, $plain_return, $caption );

__END__
