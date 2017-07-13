#!/usr/bin/perl

use lib '../lib';
use Excel::Writer::XLSX;




my $workbook = Excel::Writer::XLSX->new('../../test.xlsx');

my $worksheet = $workbook->add_worksheet('Test');

$worksheet->conditional_formatting('A1', {
    'type' => 'iconSet',
    'icons' => 'xl3TrafficLights1',
    'red_value' => 80,
    'max_gte' => 0,
    'mid_gte' => 0,
    'show_value' => 0
});
