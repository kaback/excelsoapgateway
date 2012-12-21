package ExcelSoapGateway;

# See http://www.perlmonks.org/?node_id=153486

# See http://www.tek-tips.com/faqs.cfm?fid=6715

#-------------------------------------
# Client example opening existing file:
#
# use SOAP::Lite;
#
# SOAP::Lite
#  -> uri('/ExcelSoapGateway')
#  -> proxy('http://localhost:82')
#  -> openFile('c:\soap.xls', 'Tabelle1');
#
# print SOAP::Lite
#  -> uri('/ExcelSoapGateway')
#  -> proxy('http://localhost:82')
#  -> getCellValue("c1")
#  -> result;
#
# SOAP::Lite
#  -> uri('/ExcelSoapGateway')
#  -> proxy('http://localhost:82')
#  -> setCellValue("c2", "kaback was here");
# 
# SOAP::Lite
#  -> uri('/ExcelSoapGateway')
#  -> proxy('http://localhost:82')
#  -> closeFile();
#
#--------------------------------------

use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);

$Win32::OLE::Warn = 3; # Die on Errors.
# ::Warn = 2; throws the errors, but #
# expects that the programmer deals  #

my $Excel;
my $Book;
my $Sheet;

#--------------------
# Search $string in first column, set focus to cell where
# $string has been found
# -------------------
sub findString {
	my ($self, $string) = @_;
}

#--------------------
# open existing $excelfile, activate
# existing $worksheet
# -------------------
sub openFile {
	my ($self, $excelfile, $worksheet) = @_;
	print "openFile(", $excelfile, ", ", $worksheet, ")\n";

    $Excel = Win32::OLE->GetActiveObject('Excel.Application')
        || Win32::OLE->new('Excel.Application', 'Quit');

    $Excel->{DisplayAlerts}=0; 
    $Excel->{Visible}=1;
	
	$Book = $Excel->Workbooks->Open($excelfile);
	$Sheet = $Book->Worksheets($worksheet);
	

	return $excelfile;
}

#--------------------
# close file (unsaved changes will be lost so far)
# -------------------

sub closeFile {
	print "closeFile()\n";

	$Book->SaveAs($excelfile);
    $Excel->{Visible}=0;
	$Excel->Quit;
}

#--------------------
# get value of $cell
# -------------------
sub getCellValue {
	my ($self, $cell) = @_;
	print "getCellValue(", $cell, ")\n";

	return $Sheet->Range($cell)->{Value};
}

#--------------------
# set value of $cell
# -------------------
sub setCellValue {
	my ($self, $cell, $value) = @_;
	print "setCellValue(", $cell, ", ", "$value)\n";
    $Sheet->Range($cell)->{Value} = $value;
	$Book->SaveAs($excelfile);
	return $value;
}

1;

