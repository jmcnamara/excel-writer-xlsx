#!/usr/bin/perl -w

########################################################################
#
# The tests below are a subset of tests taken from XML::Writer to ensure
# XMLsimpleWriter complies with the basic functionality of that module.
#
# I'm maintaining the same structure and methodology so that I can
# borrow more tests if required. Original documentation and copyrights
# below.
#
# John McNamara 2011.
#


########################################################################
# test.pl - test script for XML::Writer module.
# Copyright (c) 1999 by Megginson Technologies.
# Copyright (c) 2003 Ed Avis <ed@membled.com>
# Copyright (c) 2004-2010 Joseph Walton <joe@kafsemo.org>
# Redistribution and use in source and compiled forms, with or without
# modification, are permitted under any circumstances.  No warranty.
########################################################################

# Before 'make install' is performed this script should be runnable with
# 'make test'. After 'make install' it should work as 'perl 01_main.t'

use strict;

use Errno;

use Test::More(tests => 11);


# Catch warnings
my $warning;

$SIG{__WARN__} = sub {
	($warning) = @_ unless ($warning);
};

sub wasNoWarning($)
{
	my ($reason) = @_;

	if (!ok(!$warning, $reason)) {
		diag($warning);
	}
}

# Constants for Unicode support
my $unicodeSkipMessage = 'Unicode only supported with Perl >= 5.8.1';

sub isUnicodeSupported()
{
	return $] >= 5.008001;
}

require Excel::Writer::XLSX::Package::XMLwriterSimple;

TEST: {

	wasNoWarning('Load without warnings');
}

use IO::File;

# The objest that will be used for testing.
my $w;

my $outputFile = IO::File->new_tmpfile or die "Unable to create temporary file: $!";

# Fetch the current contents of the scratch file as a scalar
sub getBufStr()
{
	local($/);
	binmode($outputFile, ':bytes') if isUnicodeSupported();
	$outputFile->seek(0, 0);
	return <$outputFile>;
}

# Set up the environment to run a test.
sub initEnv(@)
{
	my (%args) = @_;

	# Reset the scratch file
	$outputFile->seek(0, 0);
	$outputFile->truncate(0);
	binmode($outputFile, ':raw') if $] >= 5.006;

	# Overwrite OUTPUT so it goes to the scratch file
	$args{'OUTPUT'} = $outputFile unless(defined($args{'OUTPUT'}));

	# Set NAMESPACES, unless it's present
	$args{'NAMESPACES'} = 1 unless(defined($args{'NAMESPACES'}));

	undef($warning);
	$w = new Excel::Writer::XLSX::Package::XMLwriterSimple($outputFile) || die "Cannot create XML writer";
}

#
# Check the results in the temporary output file.
#
# $expected - the exact output expected
#
sub checkResult($$)
{
	my ($expected, $explanation) = (@_);

	my $actual = getBufStr();

	if ($expected eq $actual) {
		ok(1, $explanation);
	} else {
		my @e = split(/\n/, $expected);
		my @a = split(/\n/, $actual);

		if (@e + @a == 2) {
			is(getBufStr(), $expected, $explanation);
		} else {
			if (eval {require Algorithm::Diff;}) {
				fail($explanation);

				Algorithm::Diff::traverse_sequences( \@e, \@a, {
					MATCH => sub { diag(" $e[$_[0]]\n"); },
					DISCARD_A => sub { diag("-$e[$_[0]]\n"); },
					DISCARD_B => sub { diag("+$a[$_[1]]\n"); }
				});
			} else {
				fail($explanation);
				diag("         got: '$actual'\n");
				diag("    expected: '$expected'\n");
			}
		}
	}

	wasNoWarning('(no warnings)');
}

#
# Expect an error of some sort, and check that the message matches.
#
# $pattern - a regular expression that must match the error message
# $value - the return value from an eval{} block
#
sub expectError($$) {
	my ($pattern, $value) = (@_);
	if (!ok((!defined($value) and ($@ =~ $pattern)), "Error expected: $pattern"))
	{
		diag('Actual error:');
		if ($@) {
			diag($@);
		} else {
			diag('(no error)');
			diag(getBufStr());
		}
	}
}

# Empty element tag.
TEST: {
	initEnv();
	$w->emptyTag("foo");
	$w->end();
	checkResult("<foo />\n", 'An empty element tag');
};

# Empty element tag with XML decl.
TEST: {
	initEnv();
	$w->xmlDecl();
	$w->emptyTag("foo");
	$w->end();
	checkResult(<<"EOS", 'Empty element tag with XML declaration');
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<foo />
EOS
};

# Start/end tag.
TEST: {
	initEnv();
	$w->startTag("foo");
	$w->endTag("foo");
	$w->end();
	checkResult("<foo></foo>\n", 'A separate start and end tag');
};

# Attributes
TEST: {
	initEnv();
	$w->emptyTag("foo", "x" => "1>2");
	$w->end();
	checkResult("<foo x=\"1&gt;2\" />\n", 'Simple attributes');
};

# Character data
TEST: {
	initEnv();
	$w->startTag("foo");
	$w->characters("<tag>&amp;</tag>");
	$w->endTag("foo");
	$w->end();
	checkResult("<foo>&lt;tag&gt;&amp;amp;&lt;/tag&gt;</foo>\n", 'Escaped character data');
};

__END__
