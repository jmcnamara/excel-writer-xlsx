#!/usr/bin/perl -w

#######################################################################
#
# A simple example of how to use the Excel::Writer::XLSX module to
# add all shapes (as currently implemented) to an Excel xlsx file.
#
# reverse('©'), May 2012, John McNamara, jmcnamara@cpan.org
#

use Excel::Writer::XLSX;

my $workbook  = Excel::Writer::XLSX->new( 'shape_all.xlsx' );

my ($worksheet, $last_sheet, $shape, $r) = (0, '', '', undef, 0);
while (<DATA>) {    
    chomp;
    next unless m/^\w/;             # Skip blank lines and comments
    my ($sheet, $name, $yn, $comment) = split(/\t/, $_);
    if ($last_sheet ne $sheet) {
        $worksheet = $workbook->add_worksheet($sheet);
        $r = 2;
    }
    $last_sheet = $sheet;
    if ($yn) {
        $shape = $workbook->add_shape( type => $name, text=>$name, width=>90, height =>90);
    } else {
        $shape = $workbook->add_shape( type => 'rect', text=>"$name (future)", width=>90, height =>90);
    }
    $worksheet->insert_shape($r, 2, $shape, 0, 0);
    $r += 5;
}

__END__
#Category	type	Implemented	Comments
Basic	rect	1	
Basic	parallelogram	1	
Basic	trash_can	0	?custom geometry (future)
Basic	diamond	1	
Basic	roundRect	1	
Basic	octagon	1	
Basic	triangle	1	
Basic	rtTriangle	1	
Basic	ellipse	1	
Basic	hexagon	1	
Basic	plus	1	
Basic	pentagon	1	
Basic	can	1	
Basic	cube	1	
Basic	bevel	1	
Basic	foldedCorner	1	
Basic	smileyFace	1	
Basic	Ghostbusters	0	custom geometry (future)
Basic	doughnut	0	custom geometry (future)
Basic	heart	0	custom geometry (future)
Basic	lightning	0	custom geometry (future)
Basic	rainbow	0	custom geometry (future)
Basic	bracket	0	custom geometry (future)
Basic	brace	0	custom geometry (future)
Basic	plaque	1	
Basic	leftBracket	1	
Basic	rightBracket	1	
Basic	leftBrace	1	
Basic	rightBrace	1	
			
Connector	line	1	3 Types, different endings
Connector	Bezier_curve	0	Bezier curve (future)
Connector	path	0	 (future)
Connector	bezier_path	0	 (future)
Connector	straightConnector1	1	3 Types, different endings
Connector	bentConnector3	1	3 Types, different endings
Connector	curvedConnector3	1	3 Types, different endings
			
Arrow	rightArrow	1	
Arrow	leftArrow	1	
Arrow	upArrow	1	
Arrow	downArrow	1	
Arrow	leftRightArrow	1	
Arrow	upDownArrow	1	
Arrow	4 way arrow	0	custom geometry (future)
Arrow	3 way arrow	0	custom geometry (future)
Arrow	curvedRightArrow	1	
Arrow	curvedLeftArrow	1	
Arrow	curvedUpArrow	1	
Arrow	curvedDownArrow	1	
Arrow	notchedRightArrow	1	
Arrow	homePlate	1	
Arrow	chevron	1	
Arrow	rightArrowCallout	1	
Arrow	leftArrowCallout	1	
Arrow	upArrowCallout	1	
Arrow	downArrowCallout	1	
Arrow	leftRightArrowCallout	1	
Arrow	upDownArrowCallout	1	
Arrow	4 way arrow callout	0	custom geometry (future)
			
FlowChart	flowChartProcess	1	
FlowChart	flowChartAlternateProcess	1	
FlowChart	flowChartDecision	1	
FlowChart	flowChartInputOutput	1	
FlowChart	flowChartPredefinedProcess	1	
FlowChart	flowChartInternalStorage	1	
FlowChart	flowChartDocument	1	
FlowChart	flowChartMultidocument	1	
FlowChart	flowChartTerminator	1	
FlowChart	flowChartPreparation	1	
FlowChart	flowChartManualInput	1	
FlowChart	flowChartManualOperation	1	
FlowChart	flowChartConnector	1	
FlowChart	flowChartOffpageConnector	1	
FlowChart	flowChartPunchedCard	1	
FlowChart	flowChartPunchedTape	1	
FlowChart	flowChartSummingJunction	1	
FlowChart	flowChartOr	1	
FlowChart	flowChartCollate	1	
FlowChart	flowChartSort	1	
FlowChart	flowChartExtract	1	
FlowChart	flowChartMerge	1	
FlowChart	flowChartOnlineStorage	1	
FlowChart	flowChartDelay	1	
FlowChart	flowChartMagneticTape	1	
FlowChart	flowChartMagneticDisk	1	
FlowChart	flowChartMagneticDrum	1	
FlowChart	flowChartDisplay	1	
			
Star_Banner	irregularSeal1	1	
Star_Banner	irregularSeal2	1	
Star_Banner	star4	1	
Star_Banner	star5	1	
Star_Banner	star8	1	
Star_Banner	star16	1	
Star_Banner	star24	1	
Star_Banner	star32	1	
Star_Banner	ribbon2	1	
Star_Banner	ribbon	1	
Star_Banner	ellipseRibbon2	1	
Star_Banner	ellipseRibbon	1	
Star_Banner	verticalScroll	1	
Star_Banner	horizontalScroll	1	
Star_Banner	wave	1	
Star_Banner	doubleWave	1	
			
Callout	wedgeRectCallout	1	
Callout	wedgeRoundRectCallout	1	
Callout	wedgeEllipseCallout	1	
Callout	cloudCallout	1	
Callout	borderCallout1	1	
Callout	borderCallout1	1	
Callout	borderCallout2	1	
Callout	borderCallout3	1	
Callout	callout1	1	
Callout	accentCallout1	1	
Callout	accentCallout2	1	
Callout	accentCallout3	1	
Callout	callout2	1	
Callout	callout3	1	
Callout	borderCallout1	1	
Callout	accentBorderCallout1	1	
Callout	accentBorderCallout2	1	
Callout	accentBorderCallout3	1	
