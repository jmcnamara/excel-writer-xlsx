#!/usr/bin/perl -w

use ExtUtils::MakeMaker;
use strict;

require 5.010;

my %deps = (
    'XML::Writer' => 0,
);

my %resources = (
    repository  => 'http://github.com/jmcnamara/excel-xlsx-writer',
    MailingList => 'http://groups.google.com/group/spreadsheet-writeexcel',
);

my %params = (
    'AUTHOR'        => 'John McNamara (jmcnamara@cpan.org)',
    'ABSTRACT'      => 'Create a new file in the Excel 2007+ XLSX format.',
    'NAME'          => 'Excel::XLSX::Writer',
    'VERSION_FROM'  => 'lib/Excel/XLSX/Writer.pm',
    'NEEDS_LINKING' => 0,
    'dist'          => { COMPRESS => 'gzip --best', SUFFIX => 'gz' },
    'LICENSE'       => 'perl',
    'PREREQ_PM'     => {},
    'test'          => { TESTS => 't/*/*.t' }

);

if ( $ExtUtils::MakeMaker::VERSION ge '6.46' ) {
    $params{META_MERGE} = { resources => \%resources };
}

WriteMakefile(%params);
