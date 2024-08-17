#!/usr/bin/perl

use ExtUtils::MakeMaker;
use strict;
use warnings;

require 5.008002;

my %deps = (
    'Archive::Zip'     => 1.30,
    'IO::File'         => 1.14,
    'File::Temp'       => 0.19,
    'File::Path::Tiny' => 0,
);

my %resources = (
    homepage   => 'http://github.com/jmcnamara/excel-writer-xlsx',
    repository => 'http://github.com/jmcnamara/excel-writer-xlsx',
    bugtracker => 'http://github.com/jmcnamara/excel-writer-xlsx/issues',
);

my %params = (
    'AUTHOR'        => 'John McNamara (jmcnamara@cpan.org)',
    'ABSTRACT'      => 'Create a new file in the Excel 2007+ XLSX format.',
    'NAME'          => 'Excel::Writer::XLSX',
    'VERSION_FROM'  => 'lib/Excel/Writer/XLSX.pm',
    'NEEDS_LINKING' => 0,
    'PREREQ_PM'     => \%deps,
    'EXE_FILES'     => ['bin/extract_vba'],
    'test'          => { TESTS => 't/*/*.t t/*/*/*.t' },
    'dist'          => {
        TAR      => 'gtar',
        COMPRESS => 'gzip --best',
        SUFFIX   => 'gz'
    },
);

if ( $ExtUtils::MakeMaker::VERSION ge '6.46' ) {
    $params{META_MERGE} = { resources => \%resources };
}

if ( $ExtUtils::MakeMaker::VERSION ge '6.30_01' ) {
    $params{LICENSE} = 'perl';
}


WriteMakefile( %params );

###############################################################################
#
# Additional target for Excel::Writer::XLSX:
#    mydocs: make the Examples.pm doc.
#
sub MY::postamble {
    "mydocs:\n"
      . "\tperl -Mlib=lib examples/gen_examples_pod.pl ./examples "
      . "> lib/Excel/Writer/XLSX/Examples.pm\n\n"
      . "tags:\n"
      . "\trm -f TAGS\n"
      . "\tetags lib/Excel/Writer/XLSX/*.pm lib/Excel/Writer/XLSX/Package/*.pm\n";
}
