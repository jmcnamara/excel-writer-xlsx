###############################################################################
#
# Tests the output of Excel::Writer::XLSX against Excel generated files.
#
# Copyright 2000-2025, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions qw(_compare_xlsx_files _is_deep_diff);
use strict;
use warnings;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $filename     = 'properties06.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = {};


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

my $company_guid = "2096f6a2-d2f7-48be-b329-b73aaa526e5d";
my $site_id      = "cb46c030-1825-4e81-a295-151c039dbf02";
my $action_id    = "88124cf5-1340-457d-90e1-0000a9427c99";


$workbook->set_custom_property("MSIP_Label_${company_guid}_Enabled",     'true',                 'text');
$workbook->set_custom_property("MSIP_Label_${company_guid}_SetDate",     '2024-01-01T12:00:00Z', 'text');
$workbook->set_custom_property("MSIP_Label_${company_guid}_Method",      'Privileged',           'text');
$workbook->set_custom_property("MSIP_Label_${company_guid}_Name",        'Confidential',         'text');
$workbook->set_custom_property("MSIP_Label_${company_guid}_SiteId",      $site_id,               'text');
$workbook->set_custom_property("MSIP_Label_${company_guid}_ActionId",    $action_id,             'text');
$workbook->set_custom_property("MSIP_Label_${company_guid}_ContentBits", '2',                    'text');

$workbook->close();


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



