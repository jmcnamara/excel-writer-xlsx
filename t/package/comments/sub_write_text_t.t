###############################################################################
#
# Tests for Excel::Writer::XLSX::Package::Comments methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions '_new_object';
use strict;
use warnings;
use Excel::Writer::XLSX::Package::Comments;

use Test::More tests => 5;


###############################################################################
#
# Tests setup.
#
my $expected;
my $got;
my $caption;
my $comment;


###############################################################################
#
# Test the _write_text_t() method.
#
$caption  = " \tComments: _write_text_t()";
$expected = '<t>Some text</t>';

$comment = _new_object( \$got, 'Excel::Writer::XLSX::Package::Comments' );

$comment->_write_text_t( 'Some text' );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_text_t() method.
#
$caption  = " \tComments: _write_text_t()";
$expected = '<t xml:space="preserve"> Some text</t>';

$comment = _new_object( \$got, 'Excel::Writer::XLSX::Package::Comments' );

$comment->_write_text_t( ' Some text' );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_text_t() method.
#
$caption  = " \tComments: _write_text_t()";
$expected = '<t xml:space="preserve">Some text </t>';

$comment = _new_object( \$got, 'Excel::Writer::XLSX::Package::Comments' );

$comment->_write_text_t( 'Some text ' );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_text_t() method.
#
$caption  = " \tComments: _write_text_t()";
$expected = '<t xml:space="preserve"> Some text </t>';

$comment = _new_object( \$got, 'Excel::Writer::XLSX::Package::Comments' );

$comment->_write_text_t( ' Some text ' );

is( $got, $expected, $caption );


###############################################################################
#
# Test the _write_text_t() method.
#
$caption  = " \tComments: _write_text_t()";
$expected = qq(<t xml:space="preserve">Some text\n</t>);

$comment = _new_object( \$got, 'Excel::Writer::XLSX::Package::Comments' );

$comment->_write_text_t( "Some text\n" );

is( $got, $expected, $caption );

__END__


