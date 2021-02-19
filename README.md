# AutoReport

This macro has been tested in versions 6.2 and 6.7 of the Aloha point of sales system. It is used to automate daily sales and payroll percentage reports.

It is run via hotkey from the Master Sales spreadsheet after End of Day (finalizes the current data - makes it read-only) has been run. It exports and prints Aloha reports via command line (reports exported to .csv format) and then searches the .csv files for needed data. After doing any necessary math, it will put data in the appropriate position on the Master sheet and the payroll percentage spreadsheet. Finally it will email the reports to the corporate office.

It is made possible by RPT.EXE supporting input via command line (likely supported in any version of Aloha post 5.0, perhaps earlier)

RPT.EXE uses the following syntax and parameters:

1./IBERDIR declares the location of the Aloha application software folder.  This does not need to be declared if the variable already resides in the system environment.

2./DATE specifies the DOB to report.  It is based on the dated folder's label.  Winhook uses the %1 variable to specify the date.  The %1 parameter is standard syntax for referring to command-line variables from within a batch file.  Refer to document ID 5998 for more information on the %1 parameter.

3./Rz to specifies the report to print.  Replace z with a parameter listed later in this document.

4.In versions 5.2x and higher, use /LOAD "y" to specify a setting file for the Labor Report, Sales Report, and Product Mix Report, where y if the report settings file name.

5. In versions 5.2x and higher, use /Xz in place of /Rz to export the report to a file rather than print the report.  Replace z with a parameter listed later in this document.  The report is exported using the format and location information already existing in the report's export settings.  The export will not function properly unless you have previously saved these settings in Aloha Manager, which creates an .EXP file for this report.

Report Parameters
Parameter and Function
A=Labor Report.  In versions 5.0x and lower, use /NUM n to declare the Labor Report number (1-3).
 
AD=Daily Cashout Summary
 
B=Employee Break Report
 
C=Sales Report
 
CW=Weekly Sales Report
 
D=ADP Payroll Export File
 
DD=Delivery Driver Report
 
DP=Delivery Production Report
 
E=Entertainer Income Report
 
F=ReMacs Menu Item Sales Export File
 
G=Gift Certificate Tracking Report
 
H=Hourly Sales and Labor Report
 
HW=Weekly Hourly Sales and Labor Report
 
I=Menu Item Forecast Report
 
J=Tip Income Report
 
K=Comparative Server Sales Report
 
L=Scheduled Vs Actual Labor Report.  You cannot fully automate this report from a command line.
 
M=Edited Punches Report
 
N=Surcharge Report
 
O=Overtime Warning Report
 
OD=Open Drawer Report
 
P=Product Mix Report
 
PQ=Weekly Quick Count Report
 
PW=Weekly Product Mix Report
 
Q=Employee Performance Measures Report
 
R=RealWorld Payroll Export File
 
S=Landry's Sales Report
 
T=Top Item Movement Report
 
TS=Team Service Tip Split Report
 
U=Edit Deposits Report
 
V=Void Report
 
W=Server Sales Report (There is not a way to select the employees to list in this report from a command line, so printing this report from a command line produces a blank report.)
 
Y=Payment Detail Report
 
Z=Sales By Revenue Center Report
 
1=Coconut Code Payroll Export File
 
11=Overtime Forecast Report
 
2=Coconut Code Sales Mix Export File
 
3=Coconut Code Daily Sales Export File
 
4=Back-of-House User Security
 
5=Back-of-House Security Levels
 
6=Detailed Access Levels
 
7=Front-of-House Cash Owed Report
 
8=Speed of Service Report
 
9=Tiered Tax Report
 
An example

The following command line exports a Sales (Cash) Report for 1/7/2000 in versions 5.2x and higher using the 'default' settings:

%IBERDIR%\BIN\RPT.EXE /DATE 20000107 /XC /LOAD "DEFAULT.SLS.SET"

Reference: https://www.tek-tips.com/viewthread.cfm?qid=1174178


