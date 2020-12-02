###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_compare_xlsx_files _is_deep_diff);
use strict;
use warnings;

use MIME::Base64;
use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $filename     = 'image51.xlsx'; #NB: reuse 51 xlsx
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with image(s) 
# from binary reference.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();


my $red_filename  = $dir . 'images/red.png';
my $red2_filename  = $dir . 'images/red2.png';
my $red   = slurp_file($red_filename);
my $red2  = slurp_file($red2_filename);



$worksheet->insert_image( 'E9',  \$red,  {url => 'https://duckduckgo.com/?q=1'});
$worksheet->insert_image( 'E13', \$red2, {url => 'https://duckduckgo.com/?q=2'});

$workbook->close();

sub slurp_file{
	my $filename = shift;
	my $string;
	open my $fh, '<', $filename or die "Couldn't open file: $!";
	local $/ = undef;
	binmode $fh;
	$string = <$fh>;
	close $fh;
	return $string;
}	
###############################################################################
#
# Compare the generated and existing Excel files.
#

my ( $got, $expected, $caption ) = _compare_xlsx_files(

    $got_filename,
    $exp_filename,
    $ignore_members,
    $ignore_elements,
);

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Cleanup.
#
unlink $got_filename;

__END__



