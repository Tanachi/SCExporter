# SCExporter
Exports story data from Sharpcloud into 2 spreadsheets

# How to install

Add References

Project -> Add References
System.Configuration
System.Drawing

COM

Microsoft Excel 16.0 Object Library

Install Packages

Tools -> Nuget Package Manager -> Package Manager Console
Install-Package Newtonsoft.Json
Install-Package SharpCloud.ClientAPI -Version 1.0.18

# Issues:
If the program crashes, a instance of excel will still be up. Must close in task manager.

Downloads Image files and reference items. Might cause problems if files are big.

Program Crashes if you have any of the spreadsheets created from this program open during runtime.
