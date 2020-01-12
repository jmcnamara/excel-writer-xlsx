###############################################################################
#
# Tests for Excel::Writer::XLSX::Worksheet methods.
#
# Copyright 2000-2020, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::Writer::XLSX::Worksheet;

use Test::More tests => 18;


###############################################################################
#
# Tests setup.
#
my $got;
my $expected;
my $caption;
my $row;
my $col;
my $worksheet;


###############################################################################
#
# 1. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 0;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:16', '17:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 2. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 1;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:15', '16:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 3. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 2;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:14', '15:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 4. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 3;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:13', '14:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 5. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 4;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:12', '13:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 6. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 5;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:11', '12:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 7. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 6;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:10', '11:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 8. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 7;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:9', '10:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 9. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 8;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:8', '9:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 10. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 9;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:7', '8:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 11. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 10;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:6', '7:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 12. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 11;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:5', '6:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 13. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 12;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:4', '5:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 14. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 13;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:3', '4:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 15. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 14;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:2', '3:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 16. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 15;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ '1:1', '2:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 17. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 16;
$col       = 0;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ undef, '1:16', '17:17' ];

is_deeply( $got, $expected, $caption );


###############################################################################
#
# 18. Test _calculate_spans() method for range ($row, $col), ($row+16, $col+16).
#
$row       = 16;
$col       = 1;
$caption   = " \tWorksheet: _calculate_spans()";
$worksheet = new Excel::Writer::XLSX::Worksheet;

for ($row .. $row +16) {
    $worksheet->write($row++, $col++, 1);
}
$worksheet->_calculate_spans();

$got = $worksheet->{_row_spans};
$expected = [ undef, '2:17', '18:18' ];

is_deeply( $got, $expected, $caption );


__END__
