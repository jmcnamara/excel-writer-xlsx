###############################################################################
#
# Tests for Excel::Writer::XLSX::Utility.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Utility qw(xl_cell_to_rowcol);

use Test::More tests => 12;

###############################################################################
#
# Tests setup.
#
my @got;
my @expected;
my $caption;
my $cell;
my $row;
my $col;

# Create a test case for a range of the Excel 2007 columns.
$cell = 'A';
for my $i ( 0 .. 300 ) {
    push @expected, [ $i, $i, 0, 0, $cell . ( $i + 1 ) ];
    $cell++;
}

$cell = 'WQK';
for my $i ( 16_000 .. 16_384 ) {
    push @expected, [ $i, $i, 0, 0, $cell . ( $i + 1 ) ];
    $cell++;
}


###############################################################################
#
# Test the xl_cell_to_rowcol method for the range of values generated above.
#
$caption = " \tUtility: xl_cell_to_rowcol()";

for my $aref ( @expected ) {
    push @got, [ xl_cell_to_rowcol( $aref->[4] ), $aref->[4] ];
}

is_deeply( \@got, \@expected, $caption );


###############################################################################
#
# Test the xl_cell_to_rowcol method with absolute references.
#
$cell     = 'A1';
@expected = ( 0, 0, 0, 0 );
$caption  = _get_caption( $cell, @expected );
@got      = xl_cell_to_rowcol($cell);
is_deeply( \@got, \@expected, $caption );

$cell     = 'A$1';
@expected = ( 0, 0, 1, 0 );
$caption  = _get_caption( $cell, @expected );
@got      = xl_cell_to_rowcol($cell);
is_deeply( \@got, \@expected, $caption );

$cell = '$A1';
@expected = ( 0, 0, 0, 1 );
$caption = _get_caption( $cell, @expected );
@got = xl_cell_to_rowcol( $cell );
is_deeply( \@got, \@expected, $caption );

$cell = '$A$1';
@expected = ( 0, 0, 1, 1 );
$caption = _get_caption( $cell, @expected );
@got = xl_cell_to_rowcol( $cell );
is_deeply( \@got, \@expected, $caption );


###############################################################################
#
# Test the xl_cell_to_rowcol examples in the docs.
#
($row, $col) = xl_cell_to_rowcol('A1');     # (0, 0)
@expected = ( 0, 0 );
$caption  = _get_caption( $cell, @expected );
@got      = xl_cell_to_rowcol($cell);
is_deeply( [ $row, $col ], \@expected, $caption );

($row, $col) = xl_cell_to_rowcol('B1');     # (0, 1)
@expected = ( 0, 1 );
$caption  = _get_caption( $cell, @expected );
@got      = xl_cell_to_rowcol($cell);
is_deeply( [ $row, $col ], \@expected, $caption );

($row, $col) = xl_cell_to_rowcol('C2');     # (1, 2)
@expected = ( 1, 2 );
$caption  = _get_caption( $cell, @expected );
@got      = xl_cell_to_rowcol($cell);
is_deeply( [ $row, $col ], \@expected, $caption );

($row, $col) = xl_cell_to_rowcol('$C2' );   # (1, 2)
@expected = ( 1, 2 );
$caption  = _get_caption( $cell, @expected );
@got      = xl_cell_to_rowcol($cell);
is_deeply( [ $row, $col ], \@expected, $caption );

($row, $col) = xl_cell_to_rowcol('C$2' );   # (1, 2)
@expected = ( 1, 2 );
$caption  = _get_caption( $cell, @expected );
@got      = xl_cell_to_rowcol($cell);
is_deeply( [ $row, $col ], \@expected, $caption );

($row, $col) = xl_cell_to_rowcol('$C$2');   # (1, 2)
@expected = ( 1, 2 );
$caption  = _get_caption( $cell, @expected );
@got      = xl_cell_to_rowcol($cell);
is_deeply( [ $row, $col ], \@expected, $caption );


###############################################################################
#
# Test error condition.
#
$cell     = '';
@expected = ( 0, 0, 0, 0 );
$caption  = _get_caption( $cell, @expected );
@got      = xl_cell_to_rowcol($cell);
is_deeply( \@got, \@expected, $caption );


###############################################################################
#
# Generate a caption for xl_cell_to_rowcol the tests.
#
sub _get_caption {

    my $cell   = shift;
    my @coords = @_;
    my $caption = " \tUtility: xl_cell_to_rowcol()";

    if ($cell) {
        $caption .=  ': ' . $cell . " -> ( " . join( ', ', @expected) . ")";
    }

    return $caption;
}

__END__


