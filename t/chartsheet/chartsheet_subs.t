###############################################################################
#
# Tests for Excel::Writer::XLSX::Chartsheet methods.
#
# reverse ('(c)'), December 2011, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_is_deep_diff);
use strict;
use warnings;
use Excel::Writer::XLSX::Chartsheet;
use Excel::Writer::XLSX::Chartsheet;

use Test::More tests => 1;


###############################################################################
#
# Compare the subroutines in Chart/Chartsheet modules.
#
my $caption = " \tChartsheet: validate subroutines.";

my @expected = _get_module_subs('Excel::Writer::XLSX::Chart');
my @got = _get_module_subs('Excel::Writer::XLSX::Chartsheet');

_is_deep_diff( \@got, \@expected, $caption );


###############################################################################
#
# Find the subroutines in Chart/Chartsheet modules.
#
sub _get_module_subs {

    no strict 'refs';

    my $module = shift;

    # Get the module functions.
    my @subs = sort keys %{"$module\::"};

    # Only return the set_ type functions.
    @subs = grep { /^[a-z]+_/ } @subs;

    # Ignore xl_ imported functions.
    @subs = grep { /^[^x][^l]/ } @subs;
}


__DATA__
