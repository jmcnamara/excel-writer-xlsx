# Excel::Writer::XLSX: Bug Reports and Pull Requests


## Reporting Bugs

Here are some tips on reporting bugs in Excel::Writer::XLSX.


### Upgrade to the latest version of the module

The bug you are reporting may already be fixed in the latest version of the
module. You can check which version of Excel::Writer::XLSX that you are using as follows:

    perl -le 'eval "require $ARGV[0]" and print $ARGV[0]->VERSION' Excel::Writer::XLSX



The [Changes](https://github.com/jmcnamara/excel-writer-xlsx/blob/master/Changes) file lists what has changed in the latest versions.


### Read the documentation

Read or search the [Excel::Writer::XLSX documentation](https://metacpan.org/pod/Excel::Writer::XLSX) to see if the issue you are encountering is already explained.

### Look at the example programs

There are [many example programs](https://metacpan.org/pod/Excel::Writer::XLSX::Examples) in the distribution. Try to identify an example program that corresponds to your query and adapt it to use as a bug report.


### Pointers for submitting a bug report

1. Describe the problem as clearly and as concisely as possible.
2. Include a sample program. This is probably the most important step. It is generally easier to describe a problem in code than in written prose.
3. The sample program should be as small as possible to demonstrate the problem. Don't copy and paste large non-relevant sections of your program.


### Sample Bug Report

A sample bug report is shown below. This format helps to analyse and respond to the bug report more quickly.

> **Issue with SOMETHING**
>
> I am using Excel::Writer::XLSX and I have encountered a problem. I
> want it to do SOMETHING but the module appears to do SOMETHING_ELSE.
>
> Here is some code that demonstrates the problem.
>
>     #!/usr/bin/perl -w
>
>     use strict;
>     use Excel::Writer::XLSX;
>
>     my $workbook  = Excel::Writer::XLSX->new("reload.xls");
>     my $worksheet = $workbook->add_worksheet();
>
>     $worksheet->write(0, 0, "Hi Excel!");
>
>     __END__
>

There is also a [bug_report.pl](https://github.com/jmcnamara/excel-writer-xlsx/blob/master/examples/bug_report.pl) program in the distro which will generate a sample report with module version numbers.

### Use the Excel::Writer::XLSX GitHub issue tracker

Submit the bug report using the [Excel::Writer::XLSX issue tracker](https://github.com/jmcnamara/excel-writer-xlsx/issues).


## Pull Requests and Contributing to Excel::Writer::XLSX

All patches and pull requests are welcome but must start with an issue tracker.


### Getting Started

1. Pull requests and new feature proposals must start with an [issue tracker](https://github.com/jmcnamara/excel-writer-xlsx/issues). This serves as the focal point for the design discussion.
2. Describe what you plan to do. If there are API changes add some code example to demonstrate them.
3. Fork the repository.
4. Run all the tests to make sure the current code work on your system using `make test`.
5. Create a feature branch for your new feature.
6. Enable Travis CI on your fork, see below.


### Enabling Travis CI via your GitHub account

Travis CI is a free Continuous Integration service that will test any code you push to GitHub with Perl 5.8, 5.10, 5.12, 5.14, 5.16 and 5.18.

See the [Travis CI Getting Started](http://about.travis-ci.org/docs/user/getting-started/) instructions.

Note there is already a `.travis.yml` file in the Excel::Writer::XLSX repo so that doesn't need to be created.


### Writing Tests

This is the most important step. Excel::Writer::XLSX has approximately 1000 tests and a 2:1 test to code ratio. Patches and pull requests for anything other than minor fixes or typos will not be merged without tests.

Use the existing tests in [the test directory](https://github.com/jmcnamara/excel-writer-xlsx/tree/master/t) for examples.

Ideally, new features should be accompanied by tests that compare Excel::Writer::XLSX output against actual Excel 2007 files. See the [regression](https://github.com/jmcnamara/excel-writer-xlsx/tree/master/t/regression) test files for examples. If you don't have access to Excel 2007 I can help you create input files for test cases.


### Code Style

Follow the general style of the surrounding code and format it with [perltidy](http://perltidy.sourceforge.net) and the following options:

    -mbl=2 -pt=0 -nola


### Running tests


Tests can be run locally as follows:

    prove -l -r t/

When you push your changes they will also be tested using [Travis CI](https://travis-ci.org/jmcnamara/excel-writer-xlsx/) as explained above.


### Documentation

If your feature requires it then write some POD documentation.


### Example programs

If applicable add an example program to the `examples` directory.


### Copyright and License

Copyright remains with the original author. Do not include additional copyright claims or Licensing requirements. GitHub and the `git` repository will record your contribution an it will be acknowledged it in the Changes file.


### Submitting the Pull Request

If your change involves several incremental `git` commits then `rebase` or `squash` them onto another branch so that the Pull Request is a single commit or a small number of logical commits.

Push your changes to GitHub and submit the Pull Request with a hash link to the to the Issue tracker that was opened above.
