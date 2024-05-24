#!/usr/bin/perl

###############################################################################
#
# Example of multi-threaded workbook creation with Excel::Writer::XLSX.
# In this example, each thread cleans up manually its temporary storage after
# creating each workbook.  The clean-up code in this example has not been
# properly tested and relying on File::Temp's cleanup process with a lock
# as in the other multithreaded examples would be safer.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use threads;

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
# Create the runnable for a thread to generate workbook files
# and destroy its temporary files in a controlled manner.
#
sub make_runnable {
    my @instructions = @_;
    sub {

      # Make the workbooks and destroy the temporary files in a thread-safe way.
        foreach (@instructions) {
            my $workbook = make_workbook(@$_);
            $workbook->{_tempdir_object}->unlink_on_destroy(0);
            File::Find::find(
                {
                    wanted => sub {
                        -f $_
                          ? unlink $File::Find::name
                          : rmdir $File::Find::name;
                    },
                    no_chdir        => 1,
                    bydepth         => 1,
                    untaint         => 1,
                    untaint_pattern => qr|^(.+)$|
                },
                $workbook->{_tempdir_object}
            );
        }

    };
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
