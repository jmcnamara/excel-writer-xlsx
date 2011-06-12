###############################################################################
#
# Tests for Excel::Writer::XLSX::Workbook methods.
#
# reverse('©'), January 2011, John McNamara, jmcnamara@cpan.org
#


use lib 't/lib';
use TestFunctions '_is_deep_diff';
use strict;
use warnings;
use Excel::Writer::XLSX::Workbook;

use Test::More tests => 1;


my $unsorted = [
    [ "Bar",                    1, q(Sheet2!$A$1)    ],
    [ "Bar",                    0, q(Sheet1!$A$1)    ],
    [ "Abc",                   -1, q(Sheet1!$A$1)    ],
    [ "Baz",                   -1, 0.98              ],
    [ "Bar",                    2, q('Sheet 3'!$A$1) ],
    [ "Foo",                   -1, q(Sheet1!$A$1)    ],
    [ "Print_Titler",          -1, q(Sheet1!$A$1)    ],
    [ "Print_Titlet",          -1, q(Sheet1!$A$1)    ],
    [ "_Fog",                  -1, q(Sheet1!$A$1)    ],
    [ "_Egg",                  -1, q(Sheet1!$A$1)    ],
    [ "_xlnm.Print_Titles",     0, q(Sheet1!$1:$1)   ],
    [ "_xlnm._FilterDatabase",  0, q(Sheet1!$G$1)    ],
    [ "aaa",                    1, q(Sheet2!$A$1)    ],
];


my $sorted = [
    [ "_Egg",                  -1, q(Sheet1!$A$1)    ],
    [ "_xlnm._FilterDatabase",  0, q(Sheet1!$G$1)    ],
    [ "_Fog",                  -1, q(Sheet1!$A$1)    ],
    [ "aaa",                    1, q(Sheet2!$A$1)    ],
    [ "Abc",                   -1, q(Sheet1!$A$1)    ],
    [ "Bar",                    2, q('Sheet 3'!$A$1) ],
    [ "Bar",                    0, q(Sheet1!$A$1)    ],
    [ "Bar",                    1, q(Sheet2!$A$1)    ],
    [ "Baz",                   -1, 0.98              ],
    [ "Foo",                   -1, q(Sheet1!$A$1)    ],
    [ "Print_Titler",          -1, q(Sheet1!$A$1)    ],
    [ "_xlnm.Print_Titles",     0, q(Sheet1!$1:$1)   ],
    [ "Print_Titlet",          -1, q(Sheet1!$A$1)    ],
];


my $got = Excel::Writer::XLSX::Workbook::_sort_defined_names($unsorted);

_is_deep_diff( $got, $sorted );


__END__
