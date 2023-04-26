###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions qw(_new_worksheet);
use strict;
use warnings;

use Test::More tests => 29;

###############################################################################
#
# Test the _encode_password() method.
#

my $temp;
my $worksheet = _new_worksheet( \$temp );


my @tests = (

    [ "password",                        "83AF" ],
    [ "This is a longer phrase",         "D14E" ],
    [ "0",                               "CE2A" ],
    [ "01",                              "CEED" ],
    [ "012",                             "CF7C" ],
    [ "0123",                            "CC4B" ],
    [ "01234",                           "CACA" ],
    [ "012345",                          "C789" ],
    [ "0123456",                         "DC88" ],
    [ "01234567",                        "EB87" ],
    [ "012345678",                       "9B86" ],
    [ "0123456789",                      "FF84" ],
    [ "01234567890",                     "FF86" ],
    [ "012345678901",                    "EF87" ],
    [ "0123456789012",                   "AF8A" ],
    [ "01234567890123",                  "EF90" ],
    [ "012345678901234",                 "EFA5" ],
    [ "0123456789012345",                "EFD0" ],
    [ "01234567890123456",               "EF09" ],
    [ "012345678901234567",              "EEB2" ],
    [ "0123456789012345678",             "ED33" ],
    [ "01234567890123456789",            "EA14" ],
    [ "012345678901234567890",           "E615" ],
    [ "0123456789012345678901",          "FE96" ],
    [ "01234567890123456789012",         "CC97" ],
    [ "012345678901234567890123",        "AA98" ],
    [ "0123456789012345678901234",       "FA98" ],
    [ "01234567890123456789012345",      "D298" ],
    [ "0123456789012345678901234567890", "D2D3" ],
);


for my $test ( @tests ) {
    my $password = $test->[0];
    my $expected = $test->[1];

    my $got = $worksheet->_encode_password( $password );

    is( $got, $expected, "Password = " . $password );
}


__END__
