#!/usr/bin/perl

use ExtUtils::MakeMaker;
use strict;
use warnings;

require 5.010;

my %deps = (
    'Archive::Zip' => 1.30,
    'IO::File'     => 0,
);

my %resources = (

    homepage    => 'http://jmcnamara.github.com/excel-writer-xlsx/',
    repository  => 'http://github.com/jmcnamara/excel-writer-xlsx',
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

###############################################################################
#
# Add 2 additional make targets for Excel::Writer::XLSX:
#    mydocs: make the Examples.pm doc.
#    test_with_xml_writer: test with XML::Writer. See t/lib/TestFunctions.pm
#
sub MY::postamble {
    "mydocs:\n"
      . "\tperl -Mblib examples/gen_examples_pod.pl examples "
      . "> lib/Excel/Writer/XLSX/Examples.pm\n\n"
      . "test_with_xml_writer:\n"
      . qq{\t_EXCEL_WRITER_XLSX_USE_XML_WRITER=1 }
      . q{PERL_DL_NONLAZY=1 }
      . q{$(FULLPERLRUN) "-MExtUtils::Command::MM" "-e" }
      . q{"test_harness($(TEST_VERBOSE), '$(INST_LIB)', '$(INST_ARCHLIB)')" }
      . q{$(TEST_FILES)} . "\n\n";
}

