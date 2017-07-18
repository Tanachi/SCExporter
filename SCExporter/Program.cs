using System;
using System.IO;
using System.Configuration;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Net;

/* 
Project -> Add References
System.Configuration
System.Drawing

COM
Microsoft Excel 16.0 Object Library

Tools -> Nuget Package Manager -> Package Manager Console
Install-Package Newtonsoft.Json
Install-Package SharpCloud.ClientAPI -Version 1.0.18
*/
// Grabs story info from Sharpcloud and converts the data into a relationship sheet and item sheet.
namespace SCExporter
{
    class Program
    {
        
        static void Main(string[] args)
        {
            var fileLocation = System.IO.Directory.GetParent
                (System.IO.Directory.GetParent(Environment.CurrentDirectory)
                .ToString()).ToString();
            // Loads user setting from config file
            var teamstoryid = ConfigurationManager.AppSettings["teamstoryid"];
            var portfolioid = ConfigurationManager.AppSettings["portfolioid"];
            var templateid= ConfigurationManager.AppSettings["templateid"];
            var userid = ConfigurationManager.AppSettings["user"];
            var passwd = ConfigurationManager.AppSettings["pass"];
            var URL = ConfigurationManager.AppSettings["URL"];
            var storyID = ConfigurationManager.AppSettings["story"];

            // Login and get story data from Sharpcloud
            var sc = new SharpCloudApi(userid, passwd, URL);
            var story = sc.LoadStory(storyID);
            var catArray = story.Categories;
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(story.ImageUri, (fileLocation +
                    "\\" + "Files" + "\\" + "story_Back.jpg"));
            }
            // Sort data into 2 different spreadsheets and writes them to disk.
            RelationshipSheet(story);
            ItemSheet(story);
            string[] paths = new string[2] {fileLocation + "\\itemFile.csv" ,
                fileLocation + "\\relationshipFile.csv"};

            MergeWorkbooks(fileLocation + "\\combine.xlsx", paths);
        }
        // Grabs all relationship data between 2 items.
        private static void RelationshipSheet(Story Story)
        {
            // Header line for relationship sheet.
            var relationshipFile = "Item 1,Item 2,Direction" + System.Environment.NewLine;


            // Parse through relationship data
            foreach (var line in Story.Relationships)
            {
                var itemOne = line.Item1.Name;
                var itemTwo = line.Item2.Name;
                var direction = line.Direction;
                var newLine = '"' + itemOne + '"' + "," + '"' + itemTwo + '"' + ","
                    + direction + System.Environment.NewLine;
                relationshipFile = relationshipFile + newLine;
            }
            //Write data to file
            File.WriteAllText(System.IO.Directory.GetParent(System.IO.Directory.GetParent(Environment.CurrentDirectory)
                .ToString()).ToString() + "\\relationshipFile.csv", relationshipFile);
            Console.WriteLine("RelationshipFile written");
        }
        //Grabs all item data with their attributes.
        private static void ItemSheet(Story Story)
        {
            // Initial Header File
            var itemFile = "Name,Description,Category,Start";
            // Grabs the attributes of the story
            var attData = Story.Attributes;
            var catData = Story.Categories;
            // Filters the default attributes from the story
            var attList = new List<SC.API.ComInterop.Models.Attribute>();
            Regex regex = new Regex(@"none|None|Sample");
            var attCount = 0;
            foreach(var att in attData)
            {
                // Checks to see if attribute header is a default attritube.
                Match match = regex.Match(att.Name);
                if (!match.Success)
                {
                    // Adds non-default attribute to the List and to the header line
                    attList.Add(att);
                    attCount++;
                    itemFile += "," + att.Name;
                }
            }
            // Adds the rest of the item line
            itemFile += "," + "Resources" + "," +
                "Tags" + "," + "Subcategory" + "," + "AttCount" + "," + "cat_color" + "," + 
                "file_path" + Environment.NewLine;
            // Goes through array of the story's item
            // goes through items based on category order
            foreach(var cat in catData)
            {
                foreach (var item in Story.Items)
                {
                    // check to see if category matches item category
                    if(item.Category.Name == cat.Name)
                    {
                        // file location for output
                        var fileLocation = System.IO.Directory.GetParent(System.IO.Directory.GetParent(Environment.CurrentDirectory).ToString()).ToString();
                        // Creates the initial line for the item 
                        var itemLine = '"' + item.Name + '"' + "," + '"' + item.Description + '"' + ","
                          + item.Category.Name + "," + item.StartDate;
                        // Adds the attributes to the item
                        foreach (var att in attList)
                        {
                            itemLine += "," + item.GetAttributeValueAsText(att);
                        }
                        // Adds the tags to the item
                        var tagLine = "";
                        foreach (var tag in item.Tags)
                        {
                            tagLine += tag.Text + "|"; 
                        }
     
                        // adds the resources to the line
                        var resLine = "";
                        foreach (var res in item.Resources)
                        {
                            resLine += res.Name + "~" + res.Url + "|";
                        }
                        resLine.ToString();
                        // adds the panels to the line
                        var panLine = "";
                        foreach ( var pan in item.Panels)
                        {
                            panLine += pan.Title + "|";
                        }
                        
                        // adds the sub category to the item
                        var subLine = "";
                        // checks to see if item has a subcategory
                        try
                        { 
                            subLine = item.SubCategory.Name;
                        }
                        catch
                        {
                            subLine = "null";
                        }
                        // adds the group of items to the line aswell as the color of the category
                        itemLine += ","  + tagLine + "," + subLine + "," +
                            cat.Color.A + "|" + cat.Color.R + "|" + cat.Color.G + "|" + cat.Color.B + ",";
                        // downloads the image for each item
                        try
                        {
                            using (WebClient client = new WebClient())
                            {
                                client.DownloadFile(item.ImageUri, (fileLocation + 
                                    "\\" + "Files" + "\\" + item.Name + ".jpg"));
                                itemLine += fileLocation +
                                    "\\" + "Files" + "\\" + item.Name + ".jpg";
                            }
                        } 
                        catch (Exception)
                        {
                            itemLine += "null";
                        }
                        
                        itemLine += System.Environment.NewLine;
                        // Adds the line to the file
                        itemFile += itemLine;
                    }
                    
                }
            }
            // Writes file to disk
            File.WriteAllText(System.IO.Directory.GetParent(System.IO.Directory.GetParent(Environment.CurrentDirectory).ToString()).ToString() + "\\itemFile.csv", itemFile);
            Console.WriteLine("ItemFile Written");
        }

        // method created by HuBeZa https://stackoverflow.com/a/32310557
        private static void MergeWorkbooks(string destinationFilePath, params string[] sourceFilePaths)
        {
            var app = new Application();
            app.DisplayAlerts = false; // No prompt when overriding

            // Create a new workbook (index=1) and open source workbooks (index=2,3,...)
            Workbook destinationWb = app.Workbooks.Add();
            foreach (var sourceFilePath in sourceFilePaths)
            {
                app.Workbooks.Add(sourceFilePath);
            }

            // Copy all worksheets
            Worksheet after = destinationWb.Worksheets[1];
            for (int wbIndex = app.Workbooks.Count; wbIndex >= 2; wbIndex--)
            {
                Workbook wb = app.Workbooks[wbIndex];
                for (int wsIndex = wb.Worksheets.Count; wsIndex >= 1; wsIndex--)
                {
                    Worksheet ws = wb.Worksheets[wsIndex];
                    ws.Copy(After: after);
                }
            }

            // Close source documents before saving destination. Otherwise, save will fail
            for (int wbIndex = 2; wbIndex <= app.Workbooks.Count; wbIndex++)
            {
                Workbook wb = app.Workbooks[wbIndex];
                wb.Close();
            }

            // Delete default worksheet
            after.Delete();

            // Save new workbook
            destinationWb.SaveAs(destinationFilePath);
            destinationWb.Close();

            app.Quit();
        }
    }
}
