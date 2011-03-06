###############################################################################
#
# Tests for Excel::Writer::XLSX::Chart methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_expected_to_aref _got_to_aref _is_deep_diff _new_object);
use strict;
use warnings;
use Excel::Writer::XLSX::Chart;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $chart = _new_object( \$got, 'Excel::Writer::XLSX::Chart' );


###############################################################################
#
# Test the _assemble_xml_file() method.
#
$caption = " \tChart: _assemble_xml_file()";

$chart->_assemble_xml_file();

$expected = _expected_to_aref();
$got      = _got_to_aref( $got );

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="en-US"/>
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$A$1:$A$5</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:order val="1"/>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$1:$B$5</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="53850880"/>
        <c:axId val="82642816"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="53850880"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:axPos val="l"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="82642816"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="82642816"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:axPos val="b"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="53850880"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:printSettings>
    <c:headerFooter/>
    <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
    <c:pageSetup/>
  </c:printSettings>
</c:chartSpace>
