#!/usr/bin/perl

use lib '../lib';
use Excel::Writer::XLSX;




my $workbook = Excel::Writer::XLSX->new('../../test.xlsx');

my $worksheet = $workbook->add_worksheet('Test');
$worksheet->write_number( 0, 0 , 95 );
$worksheet->write_number( 0, 1, 55 );
$worksheet->write_number( 2, 2, 15 );
$worksheet->write_number( 0, 3 , 5);

$worksheet->conditional_formatting( 0, 0, 5, 5, {
    'type' => 'iconSet',
    'icons' => '4TrafficLights',
    'top_value' => 80,
    'mid_value' => 40,
    'bot_value' => 10,
    'ext_value' => -1,
});
