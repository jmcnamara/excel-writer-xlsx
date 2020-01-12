###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;

use Test::More tests => 2;

###############################################################################
#
# Tests setup.
#
my $filename     = 'sub_close.xlsx';
my $dir          = 't/regression/';
my $ext_filename = $dir . $filename;
my $scalar_target;


###############################################################################
#
# Test the close() method.
#

use Excel::Writer::XLSX;

my $workbook = Excel::Writer::XLSX->new( $ext_filename );
$workbook->add_worksheet();

ok( $workbook->close(), "\tWorkbook: close(), _internal_fh" );
unlink $ext_filename;

open( my $fh, '>', \$scalar_target );
$workbook = Excel::Writer::XLSX->new( $fh );
$workbook->add_worksheet();

ok( $workbook->close(), "\tWorkbook: close(), not _internal_fh" );

__END__


