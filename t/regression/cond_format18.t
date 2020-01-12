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

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $filename     = 'cond_format18.xlsx';
my $dir          = 't/regression/';
my $got_filename = $dir . "ewx_$filename";
my $exp_filename = $dir . 'xlsx_files/' . $filename;

my $ignore_members  = [];
my $ignore_elements = { 'xl/workbook.xml' => ['<workbookView'], };


###############################################################################
#
# Test the creation of a simple Excel::Writer::XLSX file with conditional
# formatting.
#
use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( $got_filename );
my $worksheet = $workbook->add_worksheet();

$worksheet->write( 'A1', 1 );
$worksheet->write( 'A2', 2 );
$worksheet->write( 'A3', 3 );
$worksheet->write( 'A4', 4 );
$worksheet->write( 'A5', 5 );
$worksheet->write( 'A6', 6 );
$worksheet->write( 'A7', 7 );
$worksheet->write( 'A8', 8 );
$worksheet->write( 'A9', 9 );

$worksheet->write( 'A12', 75 );


$worksheet->conditional_formatting( 'A1',
    {
        type          => 'icon_set',
        icon_style    => '3_arrows',
        reverse_icons => 1,
    }
);

$worksheet->conditional_formatting( 'A2',
    {
        type         => 'icon_set',
        icon_style   => '3_flags',
        icons_only   => 1,
    }
);

$worksheet->conditional_formatting( 'A3',
    {
        type          => 'icon_set',
        icon_style    => '3_traffic_lights_rimmed',
        icons_only    => 1,
        reverse_icons => 1,
    }
);

$worksheet->conditional_formatting( 'A4',
    {
        type         => 'icon_set',
        icon_style   => '3_symbols_circled',
        icons        => [ {value => 80},
                          {value => 20},
                        ],
    }
);

$worksheet->conditional_formatting( 'A5',
    {
        type         => 'icon_set',
        icon_style   => '4_arrows',
        icons        => [ {criteria => '>'},
                          {criteria => '>'},
                          {criteria => '>'},
                        ],
    }
);

$worksheet->conditional_formatting( 'A6',
    {
        type         => 'icon_set',
        icon_style   => '4_red_to_black',
        icons        => [ {criteria => '>=', type => 'number',     value => 90},
                          {criteria => '<',  type => 'percentile', value => 50},
                          {criteria => '<=', type => 'percent',    value => 25},
                        ],
    }
);

$worksheet->conditional_formatting( 'A7',
    {
        type         => 'icon_set',
        icon_style   => '4_traffic_lights',
        icons        => [ {value => '=$A$12'} ],
    }
);

$worksheet->conditional_formatting( 'A8',
    {
        type         => 'icon_set',
        icon_style   => '5_arrows_gray',
        icons        => [ {type => 'formula', value => '=$A$12'} ],
    }
);

$worksheet->conditional_formatting( 'A9',
    {
        type          => 'icon_set',
        icon_style    => '5_quarters',
        icons         => [ { value => 70 },
                           { value => 50 },
                           { value => 30 },
                           { value => 10 },
                         ],
        reverse_icons => 1,
    }
);

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
