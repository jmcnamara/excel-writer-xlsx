#!/usr/bin/perl

###############################################################################
#
# Example of multi-threaded workbook creation with Excel::Writer::XLSX.
# In this example, each thread waits for a process-wide lock to do its own
# clean-up after creating each workbook.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use threads;
use threads::shared;

###############################################################################
#
# Declare the shared variable $cleanup_lock_counter.
#
my $cleanup_lock_counter = 0;
share $cleanup_lock_counter;

###############################################################################
#
# Generate workbooks in multiple concurrent threads.
#
my %operators = (
    addition       => '+',
    subtraction    => '-',
    multiplication => '*',
    division       => '/',
    exponentiation => '^',
);
while ( my ( $operation, $operator ) = each %operators ) {
    threads->create(
        make_runnable(
            [ $operation, $operator, 400, 400 ],
            [ $operation, $operator, 420, 420 ],
        )
    );
}
$_->join foreach threads->list();

###############################################################################
#
# Create the runnable for a thread to generate workbook files.
#
sub make_runnable {
    my @instructions = @_;
    sub {
        make_workbook_and_clean_up(@$_) foreach @instructions;
    };
}

###############################################################################
#
# Wrapper around make_workbook to manage temporary file cleanup.
#
sub make_workbook_and_clean_up {
    my ( $operation, $operator, $num_rows, $num_cols ) = @_;

    {    # increment the shared variable $cleanup_lock_counter.
        lock $cleanup_lock_counter;
        ++$cleanup_lock_counter;
    }

    # Make the workbook and store the workbook object.
    my $workbookObject =
      make_workbook( $operation, $operator, $num_rows, $num_cols );

    {    # decrement the shared variable $cleanup_lock_counter.
        lock $cleanup_lock_counter;
        --$cleanup_lock_counter;
        cond_signal $cleanup_lock_counter;
    }

    # Once $cleanup_lock_counter is 0, destroy temporary files.
    lock $cleanup_lock_counter;
    cond_wait $cleanup_lock_counter while $cleanup_lock_counter > 0;
    $workbookObject->DESTROY;
    cond_signal $cleanup_lock_counter;

}

###############################################################################
#
# Create a Excel::Writer::XLSX file.
#
sub make_workbook {

    my ( $operation, $operator, $num_rows, $num_cols ) = @_;
    my $long_name = "Table of $operation ($num_rows by $num_cols)";
    my $workbook  = Excel::Writer::XLSX->new("$long_name.xlsx") or return;
    my $worksheet = $workbook->add_worksheet( ucfirst($operation) );
    $worksheet->hide_gridlines(2);
    my $title_format =
      $workbook->add_format( size => 15, bold => 1, num_format => '@' );
    my $header_format = $workbook->add_format( bg_color => 46 );
    $worksheet->write( 0, 0, $long_name, $title_format );
    $worksheet->write( 2, 0, $operator,  $header_format );

    $worksheet->write( $_ + 2, 0,  $_, $header_format ) foreach 1 .. $num_rows;
    $worksheet->write( 2,      $_, $_, $header_format ) foreach 1 .. $num_cols;

    foreach my $row ( 3 .. ( 2 + $num_rows ) ) {
        foreach my $col ( 1 .. $num_cols ) {
            $worksheet->write( $row, $col,
                    '='
                  . xl_rowcol_to_cell( $row, 0, 0, 1 )
                  . $operator
                  . xl_rowcol_to_cell( 2, $col, 1, 0 ) );
        }
    }

    # Generate file.
    $workbook->close();

    # Return workbook object to caller, to control clean-up of temporary files.
    $workbook;

}

__END__
