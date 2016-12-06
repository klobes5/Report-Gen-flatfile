#!/usr/local/bin/perl
# Title: JCTRL_TPReport.pl
# Authors: Kevin Anthony Smith
# Owner: SDO
# Created: 03/01/2015
# Usage:
#   JCI Custom Partner Report
#   - Pulls tables from Teambook report
#   - Add PreXref table
#   - Add InboundRouting table
#   - Add DCTS Contacts
#	- Zips and archives the Report
#   - Emails Report to a given address using the "--email" flag
# Change Management:
#  ;;;$Header: /clients/fmgr_ecsc/home/fmgr/ecsc/data/JCTRL_Johnson_Controls/SCRIPTS/RCS/JCTRL_TPReport.pl,v 1.2 2015/06/22 14:24:10 fmbowner Exp $
# 2015/05/27
#       Initial Deployment
# 2015/07/20
#       Add DCTS contact information sheet
#       
BEGIN {
  my $HOME;
  $HOME = $ENV{'HOME'};
  push @INC,"$HOME/scr/perl_lib";
}

use strict;
use OLE::Storage_Lite;
use Spreadsheet::WriteExcel;
#use Excel::Writer::XLSX;	#XLSX file output is non-functional - all related lines are commented out.
use Excel::Writer::XLSX::Utility;
use Time::Local;
use Getopt::Long;
use DBI qw(:sql_types);
use Carp;
no warnings 'uninitialized';


my $HOME = $ENV{'HOME'};
my $OT_DIR = $ENV{'OT_DIR'};
my $clientID = "JCTRL";
my ($email, $dctsID);
my $subject = "'Johnson Controls Trading Partner Report'"; #message subject can be edited here.
#Grab email from command line argument
GetOptions (
			"email=s" => \$email,
			"dctsID=s" => \$dctsID,
		   );
do { print "\nERROR: You must provide an EMAIL using the --email option\n\n"; exit(1); } unless defined $email;
do { print "\nERROR: You must provide a DCTS ID using the --dctsID option\n\n"; exit(1); } unless defined $dctsID;

print "\nA copy of the TP Report will be sent to $email\n\n";

## Create a new Excel workbook
my $date = `date +"%m%d%Y%H%M%S"`;
chomp($date);
#my $xlsxfile = "${OT_DIR}/archive/ECSC/JCTRL_TPReport.${date}.$$.xlsx";
my $xlsfile = "${OT_DIR}/archive/ECSC/JCTRL_TPReport.${date}.$$.xls";
print "Creating workbook $xlsfile\n";	
#my $workbook  = Excel::Writer::XLSX->new("$xlsxfile");     
my $workbook  = Spreadsheet::WriteExcel->new("${xlsfile}");

# Set bold format (to be used for headings)
my $bold = $workbook->add_format();
$bold->set_bold();

## Generate trade guide extract
my $tmp_export = "${OT_DIR}/tmp/JCTRL_TPReport.tmp.${date}.$$";
print "\nStarting export - ${tmp_export}\n";
#system("cd $OT_DIR;SC_otstdump > ${tmp_export}");  ## uncomment this for live mode
system("cp $OT_DIR/otstdump.out ${tmp_export}");

## Apply formatting to export
print "\nApply formatting with SC_CustPortalFmt\n";
system("SC_CustPortalFmt $tmp_export ${clientID}");

## Generate partner reports
foreach my $datatype ("x12", "eft", "a2a", "ana") {
	print "\nProcessing ${datatype}...\n";
    system("SC_CustPortalCli -s ${datatype} -u EXPUPDATE -t T ${tmp_export}.${datatype}.txt");

	## Write to spreadsheet
	if ( `wc -l ${tmp_export}.${datatype}.txt.${clientID}` gt 1 ) {
		print "\n - Writing worksheet ${datatype}\n";
		&convertCSVtoExcel($workbook, ${datatype}, "${tmp_export}.${datatype}.txt.${clientID}");
	} else {
		print " - No ${datatype} data found\n";
	}
}

## VDA data
if ( `wc -l ${tmp_export}.vda.txt` gt 1 ) {
	print " - Writing worksheet vda\n";
	system("sed -i '1i${clientID}|HierRptDsc1|ScrnID|SndCode|RcvCode|SndTestCode|RcvTestCode|HierRptDsc2|CheckCtlNo|CheckInt|MsgCommID|MsgCompanyCode|MsgCompanyName|MsgInAttName|MsgInProdName|MsgInTestName|MsgLstRcvdCtlNo|MsgLstSentCtlNo|MsgOutAccName|MsgOutMdlName|MsgStd|MsgTestIndicator|MsgType|MsgXRef' ${tmp_export}.vda.txt")
	&convertCSVtoExcel($workbook, 'vda', "${tmp_export}.vda.txt");
} else {
	print " - No vda data found\n";
}

## Generate tables

## Add Routing table worksheet
my $routing_worksheet = $workbook->add_worksheet( "Routing Table" );
$routing_worksheet->keep_leading_zeros();
$routing_worksheet->set_tab_color( 'gray' );

## Add routing table heading
$routing_worksheet->write_string(0, 0, 'Lookup Key', $bold);
$routing_worksheet->write_string(0, 1, 'Plant Code', $bold);
$routing_worksheet->write_string(0, 2, 'Authorization ID', $bold);
$routing_worksheet->write_string(0, 3, 'MQ Receiver', $bold);
$routing_worksheet->write_string(0, 4, 'JCI Destination', $bold);
$routing_worksheet->write_string(0, 5, 'Doc Type', $bold);
$routing_worksheet->write_string(0, 6, 'XML Root Tag', $bold);
$routing_worksheet->write_string(0, 7, 'JCI Internal ID', $bold);
$routing_worksheet->write_string(0, 8, 'Business Unit', $bold);
$routing_worksheet->write_string(0, 9, 'Region', $bold);
$routing_worksheet->write_string(0, 10, 'Format Type', $bold);
$routing_worksheet->write_string(0, 11, 'JCI Directory', $bold);
$routing_worksheet->write_string(0, 12, 'TP Name', $bold);


## Add PreXref table worksheet
my $prexref_worksheet = $workbook->add_worksheet( "PreXref Lookup" );
$prexref_worksheet->keep_leading_zeros();
$prexref_worksheet->set_tab_color( 'gray' );

## Add PreXref table heading
$prexref_worksheet->write_string(0, 0, 'Lookup Key', $bold);
$prexref_worksheet->write_string(0, 1, 'Lookup Value', $bold);


#set base row / column for spreadsheet
my ($routing_row, $prexref_row) = 1;
my ($routing_col, $prexref_col) = 0;

# Open the database export
open (EXPFILE, $tmp_export) or die "$tmp_export: $!";

while (my $dbline = <EXPFILE>){
	chomp $dbline;

	## Find InboundRouting entries
	if ($dbline =~ /\|InboundRouting_PROD\|C\|/) {
		$routing_col = 0;

		##Remove trailing quote
		$dbline =~ s/\"$//;

		##Split the values in the database
		my @fields = split(/" "/, $dbline);
		my @key = split(/\|/, $fields[0]);
		my @val = split(/~/, $fields[1]);

		##Write out to the excel file
		$routing_worksheet->write_string($routing_row, $routing_col, $key[5]);
		$routing_col++;
		foreach my $value (@val) {
			$routing_worksheet->write_string($routing_row, $routing_col, $value);
			$routing_col++;
		}
		$routing_row++;
	}

	if ($dbline =~ /\|PreXrefLookup\|C\|/) {

		##Remove trailing quote
		$dbline =~ s/\"$//;

		##Split the values in the database
		my @fields = split(/" "/, $dbline);
		my @key = split(/\|/, $fields[0]);

		##Write out to the excel file
		$prexref_worksheet->write_string($prexref_row, 0, $key[5]);
		$prexref_worksheet->write_string($prexref_row, 1, $fields[1]);

		$prexref_row++;
	}
}

## DCTS Extract

print "Gathering DCTS information\n";

## Assign the database login
my $ENTDBLOGON = $ENV{'ENTDBLOGON'};

## Split up the login... should be in the format <login>/<password>@<database>
if ( $ENTDBLOGON !~ /(.+)\/(.+)@(.+)/ ) {
    ## attempt to split entdblogon into id($1) / pw($2) / sid($3) failed
    print STDERR "Can not connect to Oracle because environment variable ENTDBLOGON is not set correctly\n";
    exit 1;
}
my $Oracle_User = "$1";
my $Oracle_Pass = "$2";
my $Oracle_SID  = "$3";

## Connect to the database
my $dbh = DBI->connect( 'dbi:Oracle:' . $Oracle_SID,
					$Oracle_User,
					$Oracle_Pass) or croak "FATAL ERROR : Database connection not made: $DBI::errstr"; 

## Get the client ID from the db
my $client_select = qq{ SELECT group_id FROM glct_group WHERE group_key = '$dctsID' };
my $prep_sql = $dbh->prepare($client_select);
$prep_sql->execute();
my $parent_ID;
$prep_sql->bind_columns( undef, \$parent_ID);
$prep_sql->fetch;

## Get all the DCTS information for the client	 
my $dcts_sql = qq{ SELECT 
						e.EDI_ADDR, 
						e.EDI_QUALIFIER,
						e.EDI_ADDR_DESC,
						e.EDI_ADDR_TYPE,
						e.ROUTING_ADDR_ONLY,
						e.IC_EDI_TYPE,
						c.FIRST_NAME, 
						c.LAST_NAME, 
						c.PHONE_NUMBER, 
						c.EMAIL_ADDRESS
					FROM 
						glct_contact c,
						glct_group   g,
						edi_address  e
					WHERE 
						c.group_id = g.group_id
						and g.parent_group_id = '$parent_ID'
						and e.group_id = g.group_id
					ORDER BY 
						e.EDI_ADDR_DESC ASC,
						e.EDI_ADDR_TYPE ASC,
						e.EDI_ADDR
					};
$prep_sql = $dbh->prepare($dcts_sql);
$prep_sql->execute();
my($EDI_ADDR, $EDI_QUALIFIER, $EDI_ADDR_DESC, $EDI_ADDR_TYPE, $ROUTING_ADDR_ONLY, $IC_EDI_TYPE, $FIRST_NAME, $LAST_NAME, $PHONE_NUMBER, $EMAIL_ADDRESS);
$prep_sql->bind_columns( undef, \$EDI_ADDR, \$EDI_QUALIFIER, \$EDI_ADDR_DESC, \$EDI_ADDR_TYPE, \$ROUTING_ADDR_ONLY, \$IC_EDI_TYPE, \$FIRST_NAME, \$LAST_NAME, \$PHONE_NUMBER, \$EMAIL_ADDRESS);

## Add Contact table worksheet
my $contact_worksheet = $workbook->add_worksheet( "Partner Contact Info" );
$contact_worksheet->keep_leading_zeros();
$contact_worksheet->set_tab_color( 'gray' );

## Add Contact table heading
$contact_worksheet->write_string(0, 0, 'NAME', $bold);
$contact_worksheet->write_string(0, 1, 'ADDRESS', $bold);
$contact_worksheet->write_string(0, 2, 'QUAL', $bold);
$contact_worksheet->write_string(0, 3, 'ADDRESS TYPE', $bold);
$contact_worksheet->write_string(0, 4, 'DATA TYPE', $bold);
$contact_worksheet->write_string(0, 5, 'FIRST NAME', $bold);
$contact_worksheet->write_string(0, 6, 'LAST NAME', $bold);
$contact_worksheet->write_string(0, 7, 'PHONE NUMBER', $bold);
$contact_worksheet->write_string(0, 8, 'EMAIL ADDRESS', $bold);

#set base row / column for spreadsheet
my $contact_row = 1;

while ( $prep_sql->fetch ) {

	my $data_type;

	if ( "Y" eq $ROUTING_ADDR_ONLY ) {
		$data_type = "ROUTING_ONLY";
	} else {
		$data_type = $IC_EDI_TYPE;
	}

	##Write out to the excel file
	$contact_worksheet->write_string($contact_row, 0, $EDI_ADDR_DESC);
	$contact_worksheet->write_string($contact_row, 1, $EDI_ADDR);
	$contact_worksheet->write_string($contact_row, 2, $EDI_QUALIFIER);
	$contact_worksheet->write_string($contact_row, 3, $EDI_ADDR_TYPE);
	$contact_worksheet->write_string($contact_row, 4, $data_type);
	$contact_worksheet->write_string($contact_row, 5, $FIRST_NAME);
	$contact_worksheet->write_string($contact_row, 6, $LAST_NAME);
	$contact_worksheet->write_string($contact_row, 7, $PHONE_NUMBER);
	$contact_worksheet->write_string($contact_row, 8, $EMAIL_ADDRESS);

	$contact_row++;

}

###### SUBROUTINES #########
##
#  Write a csv file to an excel worksheet
##
sub convertCSVtoExcel {
	my ( $workbook, $sheet_name, $csv_file ) = @_;

	# Open the Comma Separated Variable file
	open (CSVFILE, $csv_file) or die "$csv_file: $!";
	
	# Create a new Excel workbook
	print "Adding spreadsheet $sheet_name\n";
	my $sheet_uc = uc $sheet_name;
	my $worksheet = $workbook->add_worksheet( "$sheet_uc Partners" );
	$worksheet->keep_leading_zeros();
	$worksheet->set_tab_color( 'yellow' );

	# Set bold format (to be used for headings)
	my $format = $workbook->add_format();
	$format->set_bold();
	
	# Row and column are zero indexed
	my $row = 0;
	
	while (my $csvline = <CSVFILE>){
		chomp $csvline;
		my @fields = split(/\|/, $csvline);
		my $col = 0;
		foreach my $field (@fields) {

			# Bold First Row
			if ( $row eq 0 ) {
				$worksheet->write_string($row, $col, $field, $format);
			} else {
				$worksheet->write_string($row, $col, $field);
			}
			$col++;
		}
		$row++;
	 }
}
$workbook->close();
#compress, encode and send the file to an email. Must cd to directory to prevent zipping of file hierarchy with file.
print("\nZipping ${OT_DIR}/archive/ECSC/JCTRL_TPReport.${date}.$$.xls to ${OT_DIR}/archive/ECSC/JCTRL_TPReport.${date}.$$.xls.zip\n");
system("cd ${OT_DIR}/archive/ECSC/; zip JCTRL_TPReport.${date}.$$.xls.zip JCTRL_TPReport.${date}.$$.xls");
print("\nBinary encoding and mailing ${OT_DIR}/archive/ECSC/JCTRL_TPReport.${date}.$$.xls.zip to $email\n");
system("cd ${OT_DIR}/archive/ECSC; uuencode JCTRL_TPReport.${date}.$$.xls.zip JCTRL_TPReport.${date}.$$.xls.zip | mail -v -s $subject $email");
#remove the large unzipped xls file and the temporary files
print("\nRemoving $xlsfile and temporary files located at ${tmp_export}\n");
system("rm -f $xlsfile");
system("rm -f ${tmp_export}*");
print("\nA copy of JCTRL_TPReport.${date}.$$.xls.zip has been archived at ${OT_DIR}/archive/ECSC/\n");
exit 0;
