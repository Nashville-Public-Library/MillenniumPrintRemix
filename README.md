# MillenniumPrintRemix
For libraries using the Innovative Interfaces, Inc. Millennium Integrated Library System. Remix Millennium print jobs that are simple text, e.g., order record print

millenniumOrderPrintRemix Installation

INSTALL JSCRIPT 
1.	copy millenniumOrderPrintRemix.js from Z:\Millennium\millenniumOrderPrintRemix.js to C:\Millennium\millenniumOrderPrintRemix.js

INSTALL REDMON 1.9
1.	Information on RedMon can be found at http://pages.cs.wisc.edu/~ghost/redmon/
2.	RedMon version 1.9 supporting Windows 7, Vista and XP SP3 7 can be downloaded from http://pages.cs.wisc.edu/~ghost/gsview/download/redmon19.zip.
3.	In Downloads folder, right click redmon19.zip > Select Extract all…
4.	In Extraction Wizard, select Next
5.	Files will be extracted to this directory: C:\gs\redmon > Next
6.	Check Show extracted files > Finish
7.	Right click C:\gs\redmon\setup64.exe > Select Run as Administrator…
8.	User Account Control warning: Yes
9.	RedMon – Redirection Port Monitor > Do you want to install the RedMon redirection port monitor? Yes
10.	Installation successful dialog > Select OK
11.	Close the Windows Explorer window

CONFIGURE PRINTER PORT RP1:
1.	In Windows 7, Open “Print Management”
2.	In the Print Management left pane, navigate to Print Management > [LOCAL COMPUTER NAME] > Ports
3.	In the Print Management left pane, right click Ports
4.	Select Add Port…
5.	In Printer Ports dialog, select Redirected Port and click New Port… button
6.	Port Name: RPT1: [should be default value]
7.	In Print Management central pane, right click RPT1: and select Configure Port…
8.	In the RPT1: Properties dialog
a.	Redirect this port to the program: cscript
b.	Arguments for the program are: /nologo C:\Millennium\millenniumOrderPrintRemix.js
c.	Output: Copy stdout to printer
d.	Printer: EPSON TM-T88[…]
e.	Run: Normal
f.	Run  as User [checked]
g.	Click OK
9.	In the Print Management left pane, right click Printers
10.	Select Add Printer…
11.	In the Network Printer Installation Wizard dialog:
a.	Add a new printer using an existing port: RPT1:
b.	Use an existing printer driver on the computer: EPSON TM-T88[…]
c.	Printer Name: Millennium Order Record Printer
d.	[Do not share]
e.	Next > Next [no test print – it will fail]
12.	Close Print Management 
13.	Have staffer log in to Windows and Millennium Acquisitions
14.	Select File > Select Printer > Standard Printer > Local Printer > Millennium Order Record Printer
15.	Open an order record
16.	Test print Order record to Millennium 
