### SCExporter
Downloads the data and files from a story to a excel spreadsheet.

Creates folder of story and downloads all images and files to the location of program.cs.

### How to install from Visual Studio

Create a new C# console application with the name SCExporter.

In the project folder, replace program.cs and app.config with the ones from this repo.

Create a new folder with the name "Files" in the project folder to contain all the story's images and resources.

Add References

Project -> Add References

System.Configuration

System.Drawing

Install Packages

Tools -> Nuget Package Manager -> Package Manager Console 

Enter these lines in the console in this order.

Install-Package Newtonsoft.Json

Install-Package SharpCloud.ClientAPI

Install-Package EPPlus

Example sharpcloud Url

https://my.sharpcloud.com/html/#/story/CopythisArea/view/

Enter your Sharpcloud username, password, and story-id in the app.config file.

### Issues: 
Downloads all images and references from the story. Might take some time to finish.

Program Crashes if you have any of the spreadsheets created from this program open during runtime.
