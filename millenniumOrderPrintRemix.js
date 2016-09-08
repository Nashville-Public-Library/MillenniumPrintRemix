// millenniumOrderPrintRemix.js
// James Staub
// Nashville Public Library
// remix Millennium print jobs that are simple text
// e.g., order record print

// 20150626 Display E PRICE when PAID data is not available  
// 20150623 Fix null match on bibliographic holds
//			Fix match on LOCATION=multi
// 20150427 Order record printouts with barcodes for bib and order record number
// 20150427 Check bibliographic level holds in WebPAC. If holds >= 5, then RUSH.

// TO DO: 
// Encode all keystrokes for Barcoders in scannable barcodes
// e.g., determine ITEM LOCATIONS from ORDER FORM, FUND and LOCATIONS, and fill
// out the whole freakin' item-record-generation table on behalf of the Barcoder

/* ASSUMPTIONS
page is 80 characters wide
we need to get this done - program for Nashville, not for every conceivable customer
basic order of printout is 
	record type, record number (e.g., b1000008), Last Updated:, Created:, Revisions:
	fixed-length field table
	variable-length fields
	with bib record information first, item or order information second 
each variable-length label happens at the first position in the line
length of fixed-length variable label and value label determine width of order ficed length field column
	e.g., in order record, a column composed of LOCATION, CDATE, CLAIM, COPIES, CODE1, CODE2, CODE3, PBACK will be 9 characters wide (LOCATION = 8, plus one padding space)
	ergo, we've got to explicitly look for each variable name
we don't need to grab the full label for fixed-length values that span multiple lines
*/

// SCRIPT
var printed = "";
var newPrint = "";

do {
	printed += WScript.StdIn.ReadLine() + "\n";
}
while (!WScript.StdIn.AtEndOfStream)

// BIB RECORD NUMBER
var re = /\n(b\d{7}[\dx])/;
var matches = re.exec(printed);
var bibRecord = matches[1];
// CALL NUMBER
var re = /\nc 09\d +?(\w.+?)\n/;
var matches = re.exec(printed);
var callNumber = "";
if (matches !== null) { callNumber = matches[1] };
// TITLE
var re = /\nt 245.... (\w.+?) *?\n/;
var matches = re.exec(printed);
var title = "";
if (matches !== null) { title = matches[1] };
// ORDER RECORD NUMBER
var re = /\n *(o\d{6,7}[\dx])([\s\S]+?$)/;
var matches = re.exec(printed);
var orderRecord = matches[1];
printed = matches[2]; // remove the bib section from the source
// ORDER LOCATION
var re = /\nLOCATION (\w.+?) +?FORM/;
var matches = re.exec(printed);
var orderLocations = "";
if (matches[1] == "multi") {
	re = /\nLOCATIONS ([\s\S]+?)\n[\w]/;
	var matches = re.exec(printed);
	if (matches[1]) { 
		orderLocations = matches[1].replace(/\s/g, "");
	}
} else if (matches !== null) {
	orderLocations = matches[1];
}
// ORDER STATUS
var re = /\sSTATUS +?(\w.+?)\n/;
var matches = re.exec(printed);
var orderStatus = "";
if (matches !== null) { orderStatus = matches[1]; }
// ORDER COPIES
var re = /\nCOPIES +?(\w.+?) +?ORD NOTE/;
var matches = re.exec(printed);
var orderCopies = "";
if (matches !== null) { orderCopies = matches[1]; }
// ORDER CODE2
var re = /\nCODE2 +?(\w.+?) +?RACTION/;
var matches = re.exec(printed);
var orderCode2 = "";
if (matches !== null) { orderCode2 = matches[1];}
// ORDER RDATE
var re = /\sRDATE +?(\d\d-\d\d-\d\d\d\d)/;
var matches = re.exec(printed);
var orderRdate = "";
if (matches !== null) { orderRdate = matches[1] };
// ORDER i
var orderI = "";
var re = /\ni +?(\w[\s\S]+?)\n[^i\s]/;
var matches = re.exec(printed);
if (matches !== null) {
	orderI = matches[1];
	orderI = orderI.replace(/\ni/g, "NEWLINENEWLINE");
	orderI = orderI.replace(/\n/g, " ");
	orderI = orderI.replace(/ +/g, " ");
	orderI = orderI.replace(/NEWLINENEWLINE/g, "\n");
}
// ORDER z
var orderZ = "";
var re = /\nz +?(\w[\s\S]+?)\n[^z\s]/;
var matches = re.exec(printed);
if (matches !== null) {
	orderZ = matches[1];
	orderZ = orderZ.replace(/\nz/g, "NEWLINENEWLINE");
	orderZ = orderZ.replace(/\n/g, " ");
	orderZ = orderZ.replace(/ +/g, " ");
	orderZ = orderZ.replace(/NEWLINENEWLINE/g, "\n");
}
// RUSH : look for ORDER z "rush" OR bibliographic holds => 5
var rush = false;
if (orderZ && (/ rush/i).test(orderZ)) {
	rush = true;
}
function getText(strURL) // from https://msdn.microsoft.com/en-us/library/windows/desktop/aa384071%28v=vs.85%29.aspx
{
    var strResult;
    try
    {
        // Create the WinHTTPRequest ActiveX Object.
        var WinHttpReq = new ActiveXObject("WinHttp.WinHttpRequest.5.1");
        //  Create an HTTP request.
        var temp = WinHttpReq.Open("GET", strURL, false);
        //  Send the HTTP request.
        WinHttpReq.Send();
        //  Retrieve the response text.
        strResult = WinHttpReq.ResponseText;
    }
    catch (objError)
    {
        strResult = objError + "\n"
        strResult += "WinHTTP returned error: " + 
            (objError.number & 0xFFFF).toString() + "\n\n";
        strResult += objError.description;
    }
    //  Return the response text.
    return strResult;
}
var bibURL = "http://waldo.library.nashville.org/record=" + bibRecord.substr(0,8);
var bibPage = getText(bibURL);
var pattern = /<span class="bibHolds">(\d+) holds* on first copy returned/;
var matches = pattern.exec(bibPage);
var bibHolds = 0;
if (matches !== null) {bibHolds = Number(matches[1]);}
if (matches && bibHolds > 4) {
	rush = true;
}
// ORDER PAID PER COPY
// calculates price per copy of the first PAID line
// calls attention to CATALOGING if standing order, unpaid or partial fulfillment
var re = /\nPAID .+?\$([\d.]+) \d+ (\d+)/;
var matches = re.exec(printed);
var receivedCopies;
var orderPaidPerCopy;
if (matches !== null) {
	receivedCopies = Number(matches[2]);
	orderPaidPerCopy = Math.ceil((Number(matches[1])/Number(matches[2]))*100)/100;
	orderPaidPerCopy = "$" + orderPaidPerCopy.toString();
} else {
	var re = / (E PRICE .+?\$[\d.]+?) /;
	var matches = re.exec(printed);
	if (matches !== null) {orderPaidPerCopy = (matches[1]);}
}
if (orderStatus.substr(0,1) == "f") { 
	orderPaidPerCopy = "STANDING ORDER. CATALOGING VERIFY" + orderPaidPerCopy;
} else if (orderStatus.substr(0,1) == "o") { 
	orderPaidPerCopy = "UNPAID ORDER. CATALOGING VERIFY " + orderPaidPerCopy;
} else if (orderCopies != receivedCopies) {
	orderPaidPerCopy = "PARTIAL FULFILLMENT. CATALOGING VERIFY" + orderPaidPerCopy;		
}


// PRINT
// EPSON ESC/POS information at http://www.epsonexpert.com/Epson_Assets/ESCPOS_Commands_FAQs.pdf
newPrint += "\x1B" + "@" // Initialize printer

//newPrint += "\x1B" + "!" + "0";
if (rush) {
	newPrint += "\x1B" + "a" + "\x01"; // centered
	newPrint += "\x1D" + "!" + "\x33"; // quad-width, quad-height
	newPrint += "\x1D" + "B" + "\x01"; // reverse black and white
	newPrint += "   RUSH   \n";
	newPrint += "\x1D" + "B" + "\x00"; // cancel reverse black and white
	newPrint += "\x1D" + "!" + "\x00"; // cancel quad-width, quad-height
	newPrint += "\nHOLDS: " + bibHolds + "\n\n";
	newPrint += "\x1B" + "a" + "\x00"; // left-justified
}

newPrint += "\x1D" + "\x68" + "\x50"; // set barcode height to 80 pt
newPrint += "\x1D" + "k" + "\x45" + "\x09" + bibRecord.toUpperCase(); // CODE39 barcode, might be wrong
newPrint += "." + bibRecord + "\n"; 
newPrint += "CALL NUMBER: " + callNumber + "\n";
newPrint += "TITLE: " + title + "\n";
newPrint += "------------------------------------------\n";
newPrint += "\x1D" + "\x68" + "\x50"; // set barcode height to 80 pt
newPrint += "\x1D" + "k" + "\x45" + "\x08" + orderRecord.toUpperCase(); // CODE39 barcode, might be wrong
newPrint += "." + orderRecord + "\n";
newPrint += "------------------------------------------\n";
newPrint += "STATUS: " + orderStatus + "\n";
newPrint += "CODE2: " + orderCode2 + "\n";
newPrint += "RECEIVED DATE: " + orderRdate + "\n";
newPrint += "------------------------------------------\n";
newPrint += "COPIES: " + orderCopies + "\n";
newPrint += "LOCATIONS: " + orderLocations + "\n";
newPrint += "------------------------------------------\n";
newPrint += "ORDER I: " + orderI + "\n";
newPrint += "------------------------------------------\n";
newPrint += "ORDER Z: " + orderZ + "\n";
newPrint += "------------------------------------------\n";
newPrint += "PER ITEM PRICE: " + orderPaidPerCopy + "\n";
newPrint +=  "\n"; // necessary?
newPrint +=  "\n"; // necessary?
newPrint += "\x1D" + "V" + "\x42"; // feed, partial cut

//WScript.StdOut.WriteLine("JAMES LOOK HERE->");
WScript.StdOut.WriteLine(newPrint);
//WScript.StdOut.WriteLine("<-HERE LOOK JAMES");

/* WHAT TO LOOK FOR

BIBLIOGRAPHIC RECORD VARIABLE LENGTH
001 > ! REC INFO 
002 > " BIB INFO 
003 > _ LEADER 
004 > c CALL # c 
005 > a AUTHOR a 
006 > t TITLE Wkt 
007 > e EDITION 
008 > p PUB INFO 
009 > r DESCRIPT 
010 > s SERIES Wat 
011 > n NOTE Wt 
012 > d SUBJECT Wd 
013 > b ADD AUTHOR at 
014 > u ADD TITLE t 
015 > x CONTINUES 
016 > z CONT'D BY 
017 > w RELATED TO 
018 > o BIB UTIL # o 
019 > i STANDARD # i 
020 > l LCCN 
021 > g GOV DOC # g 
022 > h LIB HAS 
023 > k TOC DATA Wat 
024 > y MISC 
025 > 1 LOCATIONS

BIBLIOGRAPHIC RECORD FIXED LENGTH
024 > LANG or Language* 
025 > SKIP or Skip* 
026 > LOCATION or Location* 
027 > COPIES or Copies* (Most libraries do not use this field.)
028 > CAT DATE or Cat. Date* (set via @cdate trigger) 
029 > BIB LVL or Bib Level (BCODE1)* 
030 > MAT TYPE or Material Type (BCODE2)* 
031 > BCODE3 or Bib Code 3* 
089 > COUNTRY or Country* 
107 > MARCTYPE or MARC Type*

ORDER VARIABLE LENGTH
026 > ! REC INFO 
027 > # ORDER INFO 
028 > i IDENTITY 
029 > x FOR CURR 
030 > n NOTE 
031 > z INT NOTE 
032 > s SELECTOR 
033 > r REQUESTOR 
034 > q VEN ADDR 
035 > v VEN NOTE 
036 > f VEN TITL # 
037 > b PO INFO 
038 > p BLANKET PO 
039 > j TICKLER 
040 > k TICKLERLOG 
041 > 0 PAID 
042 > 1 LOCATIONS 
043 > 2 FUNDS 
044 > 3 REC COPIES 
045 > 4 LISTPRICE 
046 > 5 REOPEN DAT 
047 > 6 STATUS REP

ORDER RECORD FIXED LENGTH
001 > ACQ TYPE or Acq Type* 
002 > LOCATION or Location* 
003 > CDATE or Cat Date* 
004 > CLAIM or Claim 
005 > COPIES or Copies* 
006 > CODE1 or Order Code 1* 
007 > CODE2 or Order Code 2* 
008 > CODE3 or Order Code 3* 
009 > CODE4 or Order Code 4* 
010 > E PRICE or Est. Price* 
011 > FORM or Form* 
012 > FUND or Fund* 
013 > ODATE or Order Date* 
014 > ORD NOTE or Order Note* 
015 > ORD TYPE or Order Type* 
016 > RACTION or Recv Action* 
017 > RDATE or Recv Date* 
018 > RLOC or Recv Location* 
019 > BLOC or Billing Location* 
020 > STATUS(O) or Status* (But see NOTE below) 
021 > TLOC or Transit Location* 
022 > VENDOR or Vendor* 
023 > LANG or Language* 
032 > PAID DATE or Paid Date 
033 > INV DATE or Invoice Date 
034 > PAID AMT or Paid Amount 
100 > COUNTRY or Country* 
106 > VOLUMES or Volumes*

*/
