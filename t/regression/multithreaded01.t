###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
# (multi-threaded variant)
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_compare_xlsx_files _is_deep_diff);
use strict;
use warnings;

my ( $numThreads, $numBooksPerThread );

BEGIN {
    $numThreads        = 6;
    $numBooksPerThread = 7;
}

use Test::More;
use Excel::Writer::XLSX;

eval { require threads; require threads::shared };
if ($@) {
    plan skip_all => 'threads and threads::shared required to run these tests.';
}
else {
    plan tests => $numThreads * $numBooksPerThread;
}

###############################################################################
#
# Tests setup.
#
my $filename     = "simple01.xlsx";
my $dir          = 't/regression/';
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};

###############################################################################
#
# Declare the shared variable $cleanup_lock_counter.
#
my $cleanup_lock_counter = 0;
threads::shared::share( \$cleanup_lock_counter );

###############################################################################
#
# Generate workbooks in multiple concurrent threads.
#
foreach my $threadNo ( 1 .. $numThreads ) {
    threads->create( make_runnable( $dir . "Thread $threadNo" ) );
}
$_->join foreach threads->list();

###############################################################################
#
# Compare the generated and existing Excel files, and cleanup.
#
foreach my $threadNo ( 1 .. $numThreads ) {
    foreach my $bookNo ( 1 .. $numBooksPerThread ) {
        my $got_filename = $dir . "Thread $threadNo book $bookNo.xlsx";

        my ( $got, $expected, $caption ) = _compare_xlsx_files(

            $got_filename, $exp_filename,
            $ignore_members,
            $ignore_elements,
        );

        _is_deep_diff( $got, $expected, $caption );

        unlink $got_filename;

    }
}

###############################################################################
#
# Create the runnable for a thread to generate workbook files
# and destroy its temporary files in a controlled manner.
#
sub make_runnable {
    my ($filenamePrefix) = @_;
    sub {
        my @workbookObjects;

        # Make workbooks, maintaining the shared variable $cleanup_lock_counter.
        foreach my $bookNo ( 1 .. $numBooksPerThread ) {
            {    # increment the shared variable $cleanup_lock_counter.
                lock $cleanup_lock_counter;
                ++$cleanup_lock_counter;
            }
            push @workbookObjects,
              make_workbook("$filenamePrefix book $bookNo.xlsx");
            {    # decrement the shared variable $cleanup_lock_counter.
                lock $cleanup_lock_counter;
                --$cleanup_lock_counter;
                threads::shared::cond_signal( \$cleanup_lock_counter );
            }
        }

        # Wait for the shared variable $cleanup_lock_counter to be zero.
        lock $cleanup_lock_counter;
        threads::shared::cond_wait( \$cleanup_lock_counter )
          while $cleanup_lock_counter > 0;

        # Destruction of temporary files associated with each workbook object.
        $_->DESTROY foreach grep { defined $_; } @workbookObjects;
        threads::shared::cond_signal( \$cleanup_lock_counter );

    };
}

###############################################################################
#
# Create a simple Excel::Writer::XLSX file which should match simple01.xlsx.
#
sub make_workbook {
    my ($filename) = @_;

    my $workbook = Excel::Writer::XLSX->new($filename) or return;
    my $worksheet = $workbook->add_worksheet();

    $worksheet->write( 'A1', 'Hello' );
    $worksheet->write( 'A2', 123 );

    # Generate file.
    $workbook->close();

    # Return workbook object to caller, to control clean-up of temporary files.
    $workbook;

}

__END__
