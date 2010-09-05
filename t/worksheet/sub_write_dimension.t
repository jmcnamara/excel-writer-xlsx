###############################################################################
#
# Tests for Excel::XLSX::Writer::Worksheet methods.
#
# reverse('©'), September 2010, John McNamara, jmcnamara@cpan.org
#

use strict;
use warnings;
use Excel::XLSX::Writer;
use XML::Writer;

use Test::More tests => 10;


###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $worksheet;
my $tmp;
my $got;
my $cell_ref;


###############################################################################
#
# 1. Test the _write_dimension() method with no dimensions set.
#
$caption  = " \tWorksheet: _write_dimension(undef)";
$expected = '<dimension ref="A1" />';

$worksheet = _new_worksheet();

$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 2. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref" />);

$worksheet = _new_worksheet();
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 3. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1048576';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref" />);

$worksheet = _new_worksheet();
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 4. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'XFD1';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref" />);

$worksheet = _new_worksheet();
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 5. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'XFD1048576';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref" />);

$worksheet = _new_worksheet();
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 6. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref" />);

$worksheet = _new_worksheet();
$worksheet->write( $cell_ref, 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 7. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1:B2';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref" />);

$worksheet = _new_worksheet();
$worksheet->write( 'A1', 'some string' );
$worksheet->write( 'B2', 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 8. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1:B2';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref" />);

$worksheet = _new_worksheet();
$worksheet->write( 'B2', 'some string' );
$worksheet->write( 'A1', 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 9. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'B2:H11';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref" />);

$worksheet = _new_worksheet();
$worksheet->write( 'B2',  'some string' );
$worksheet->write( 'H11', 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# 10. Test the _write_dimension() method with dimensions set.
#
$cell_ref = 'A1:XFD1048576';
$caption  = " \tWorksheet: _write_dimension('$cell_ref')";
$expected = qq(<dimension ref="$cell_ref" />);

$worksheet = _new_worksheet();
$worksheet->write( 'A1',         'some string' );
$worksheet->write( 'XFD1048576', 'some string' );
$worksheet->_write_dimension();

is( $got, $expected, $caption );


###############################################################################
#
# Internal function used for setting up each test.
#
sub _new_worksheet {

    $got = '';
    $tmp = '';


    open my $tmp_fh, '>', \$tmp or die "Failed to open filehandle: $!";
    open my $got_fh, '>', \$got or die "Failed to open filehandle: $!";

    my $workbook  = Excel::XLSX::Writer->new( $tmp_fh );
    my $worksheet = $workbook->add_worksheet;
    my $writer    = new XML::Writer( OUTPUT => $got_fh );

    $worksheet->{_writer} = $writer;

    return $worksheet;
}


__END__
