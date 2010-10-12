#!/usr/bin/perl

use ExtUtils::MakeMaker;
use strict;
use warnings;

require 5.010;

my %deps = (
    'Archive::Zip' => 1.30,
    'IO::File'     => 0,
    'XML::Writer'  => 0.612,
);

my %resources = (
    repository  => 'http://github.com/jmcnamara/excel-xlsx-writer',
    MailingList => 'http://groups.google.com/group/spreadsheet-writeexcel',
);

my %params = (
    'AUTHOR'        => 'John McNamara (jmcnamara@cpan.org)',
    'ABSTRACT'      => 'Create a new file in the Excel 2007+ XLSX format.',
    'NAME'          => 'Excel::Writer::XLSX',
    'VERSION_FROM'  => 'lib/Excel/Writer/XLSX.pm',
    'NEEDS_LINKING' => 0,
    'dist'          => { COMPRESS => 'gzip --best', SUFFIX => 'gz' },
    'LICENSE'       => 'perl',
    'PREREQ_PM'     => \%deps,
    'test'          => { TESTS => 't/*/*.t t/*/*/*.t' }

);

if ( $ExtUtils::MakeMaker::VERSION ge '6.46' ) {
    $params{META_MERGE} = { resources => \%resources };
}

WriteMakefile(%params);

sub MY::postamble {
    "mydocs:\n"
      . "\tperl -Mblib examples/gen_examples_pod.pl examples "
      . "> lib/Excel/Writer/XLSX/Examples.pm\n"
}
