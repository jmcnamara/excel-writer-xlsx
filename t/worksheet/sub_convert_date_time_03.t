###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet date and time handling.
#
# Tests date and time second handling.
#
# Copyright 2000-2023, John McNamara, jmcnamara@cpan.org
#
# SPDX-License-Identifier: Artistic-1.0-Perl OR GPL-1.0-or-later
#

use lib 't/lib';
use TestFunctions '_new_worksheet';
use strict;
use warnings;

use Test::More tests => 104;

my $got;
my $date_time;
my $number;
my $result;
my $worksheet = _new_worksheet( \$got );

$worksheet->{_date_1904} = 0;

# Set float difference limit to half of an Excel millisecond
my $flt_delta = 0.5 / ( 24 * 60 * 60 * 1000 );


##############################################################################
#
# Float comparison function.
#
sub flt_cmp {
    return abs( $_[0] - $_[1] ) < $flt_delta;
}


##############################################################################
#
# Test the flt_cmp() function used in the other tests.
#

$date_time = '1899-12-31T00:00:00.0004';
$number    = 0;
$result    = $worksheet->convert_date_time( $date_time );
$result    = -1 unless defined $result;

# Test 1. This should pass. It is less than the float diff limit.
ok( flt_cmp( $number, $result ),
    " Testing convert_date_time: $date_time $number" );

$date_time = '1899-12-31T00:00:00.0005';
$number    = 0;
$result    = $worksheet->convert_date_time( $date_time );
$result    = -1 unless defined $result;

# Test 2. This should fail. It is equal to the float diff limit.
my $diff = !flt_cmp( $number, $result );
ok( $diff, " Testing convert_date_time:  $date_time $number" );


##############################################################################
#
# Test some false times.
#

# These should fail.
my $fail;

$date_time = '1899-12-31T24:00:00.000';
$fail      = !$worksheet->convert_date_time( $date_time );
ok( $fail, " Testing incorrect time: $date_time\tincorrect hour caught." );

$date_time = '1899-12-31T00:60:00.000';
$fail      = !$worksheet->convert_date_time( $date_time );
ok( $fail, " Testing incorrect time: $date_time\tincorrect mins caught." );

$date_time = '1899-12-31T00:00:60.000';
$fail      = !$worksheet->convert_date_time( $date_time );
ok( $fail, " Testing incorrect time: $date_time\tincorrect secs caught." );

$date_time = '1899-12-31T00:00:59.999999999999999999';
# Workaround for increased precision with longdouble perls.
$date_time .= "9" x length (1 - 1e-16);

$fail      = !$worksheet->convert_date_time( $date_time );
ok( $fail, " Testing incorrect time: $date_time\tincorrect secs caught." );


##############################################################################
#
# Test the time data generated in Excel.
#
while ( <DATA> ) {

    last if /^# stop/;    # For debugging
    next unless /\S/;     # Ignore blank lines
    next if /^#/;         # Ignore comments

    if ( /"DateTime">([^<]+)/ ) {
        my $date_time = $1;
        my $line      = <DATA>;

        if ( $line =~ /"Number">([^<]+)/ ) {
            my $number = 0 + $1;
            my $result = $worksheet->convert_date_time( $date_time );
            $result = -1 unless defined $result;

            ok( flt_cmp( $number, $result ),
                " Testing convert_date_time: $date_time  $number" )
              or diag( "difference between $number and $result\n" . "= "
                  . abs( $number - $result ) . "\n"
                  . "> $flt_delta" );
        }
    }
}


__DATA__


# Test data taken from Excel in XML format.
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T00:00:00.000</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T00:15:20.213</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">1.0650613425925924E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T00:16:48.290</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">1.1670023148148148E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T00:55:25.446</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">3.8488958333333337E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T01:02:46.891</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">4.3598275462962965E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T01:04:15.597</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">4.4624965277777782E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T01:09:40.889</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">4.8389918981481483E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T01:11:32.560</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">4.9682407407407404E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T01:30:19.169</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">6.2721863425925936E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T01:48:25.580</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">7.5296064814814809E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T02:03:31.919</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">8.5786099537037031E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T02:11:11.986</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">9.1110949074074077E-2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T02:24:37.095</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.10042934027777778</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T02:35:07.220</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.1077224537037037</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T02:45:12.109</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.11472348379629631</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T03:06:39.990</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.12962951388888888</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T03:08:08.251</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.13065105324074075</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T03:19:12.576</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.13833999999999999</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T03:29:42.574</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.14563164351851851</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T03:37:30.813</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.1510510763888889</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T04:14:38.231</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.1768313773148148</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T04:16:28.559</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.17810832175925925</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T04:17:58.222</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.17914608796296297</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T04:21:41.794</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.18173372685185185</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T04:56:35.792</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.2059698148148148</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T05:25:14.885</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.22586672453703704</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T05:26:05.724</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.22645513888888891</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T05:46:44.068</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.24078782407407406</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T05:48:01.141</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.2416798726851852</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T05:53:52.315</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.24574438657407408</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T06:14:48.580</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.26028449074074073</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T06:46:15.738</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.28212659722222222</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T07:31:20.407</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.31343063657407405</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T07:58:33.754</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.33233511574074076</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T08:07:43.130</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.33869363425925925</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T08:29:11.091</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.35360059027777774</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T09:08:15.328</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.380732962962963</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T09:30:41.781</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.39631690972222228</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T09:34:04.462</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.39866275462962958</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T09:37:23.945</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.40097158564814817</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T09:37:56.655</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.40135017361111114</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T09:45:12.230</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.40639155092592594</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T09:54:14.782</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.41267108796296298</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T09:54:22.108</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.41275587962962962</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T10:01:36.151</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.41777952546296299</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T12:09:48.602</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.50681252314814818</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T12:34:08.549</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.52371005787037039</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T12:56:06.495</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.53896406249999995</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T12:58:58.217</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.54095158564814816</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T12:59:54.263</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.54160026620370372</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T13:34:41.331</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.56575614583333333</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T13:58:28.601</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.58227547453703699</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T14:02:16.899</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.58491781249999997</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T14:36:17.444</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.60853523148148148</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T14:37:57.451</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.60969271990740748</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T14:57:42.757</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.6234115393518519</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T15:10:48.307</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.6325035532407407</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T15:14:39.890</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.63518391203703706</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T15:19:47.988</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.63874986111111109</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T16:04:24.344</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.66972620370370362</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T16:22:23.952</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.68222166666666662</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T16:29:55.999</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.6874536921296297</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T16:58:20.259</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.70717892361111112</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T17:04:02.415</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.71113906250000003</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T17:18:29.630</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.72117627314814825</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T17:47:21.323</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.74121901620370367</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T17:53:29.866</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.74548456018518516</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T17:53:41.076</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.74561430555555563</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T17:55:06.044</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.74659773148148145</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T18:14:49.151</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.760291099537037</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T18:17:45.738</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.76233493055555546</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T18:29:59.700</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.77082986111111118</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T18:33:21.233</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.77316241898148153</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T19:14:24.673</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.80167445601851861</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T19:17:12.816</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.80362055555555545</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T19:23:36.418</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.80806039351851855</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T19:46:25.908</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.82391097222222232</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T20:07:47.314</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.83874206018518516</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T20:31:37.603</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.85529633101851854</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T20:39:57.770</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.86108530092592594</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T20:50:17.067</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.86825309027777775</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T21:02:57.827</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.87705818287037041</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T21:23:05.519</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.891036099537037</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T21:34:49.572</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.89918486111111118</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T21:39:05.944</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.90215212962962965</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T21:39:18.426</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.90229659722222222</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T21:46:07.769</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.90703436342592603</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T21:57:55.662</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.91522756944444439</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T22:19:11.732</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.92999689814814823</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T22:23:51.376</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.93323351851851843</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T22:27:58.771</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.93609688657407408</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T22:43:30.392</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.94687953703703709</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T22:48:25.834</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.95029900462962968</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T22:53:51.727</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.95407091435185187</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T23:12:56.536</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.96732101851851848</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T23:15:54.109</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.96937626157407408</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T23:17:12.632</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.97028509259259266</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s22"><Data ss:Type="DateTime">1899-12-31T23:59:59.999</Data></Cell>
    <Cell ss:StyleID="s23" ss:Formula="=RC[-1]"><Data ss:Type="Number">0.99999998842592586</Data></Cell>
   </Row>



