using System;
using System.IO;
using System.Configuration;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Net;
using System.Text;
using System.Threading;
// Grabs story info from Sharpcloud and converts the data into a relationship sheet and item sheet.
namespace SCExporter
{
    class Program
    {
        class relationship
        {
            private string storyID;
            Application app;
            Worksheet relationshipSheet;
            private static void RelationshipSheet(Story Story, Worksheet relationshipSheet)
            {
                // file path variable
                var fileLocation = System.IO.Directory.GetParent
                    (System.IO.Directory.GetParent(Environment.CurrentDirectory)
                    .ToString()).ToString();

                // Header Line
                relationshipSheet.Cells[1, "A"].Value2 = "Item 1";
                relationshipSheet.Cells[1, "B"].Value2 = "Item 2";
                relationshipSheet.Cells[1, "C"].Value2 = "Direction";
                var count = 2;

                // Parse through relationship data
                foreach (var line in Story.Relationships)
                {
                    relationshipSheet.Cells[count, "A"] = line.Item1.Name;
                    relationshipSheet.Cells[count, "B"] = line.Item2.Name;
                    relationshipSheet.Cells[count, "C"] = line.Direction.ToString();
                    count++;
                }
                //Write data to file
                relationshipSheet.SaveAs(fileLocation + "\\relationshipFile.csv", XlFileFormat.xlCSVWindows);
                Console.WriteLine("Relationship file written");

            }
        }
        static void Main(string[] args)
        {
            var fileLocation = System.IO.Directory.GetParent
                (System.IO.Directory.GetParent(Environment.CurrentDirectory)
                .ToString()).ToString();
            // Loads user setting from config file
            var teamstoryid = ConfigurationManager.AppSettings["teamstoryid"];
            var portfolioid = ConfigurationManager.AppSettings["portfolioid"];
            var templateid = ConfigurationManager.AppSettings["templateid"];
            var userid = ConfigurationManager.AppSettings["user"];
            var passwd = ConfigurationManager.AppSettings["pass"];
            var URL = ConfigurationManager.AppSettings["URL"];
            var storyID = ConfigurationManager.AppSettings["story"];

            // Login and get story data from Sharpcloud
            var sc = new SharpCloudApi(userid, passwd, URL);
            var story = sc.LoadStory(storyID);
            //Create excel application and create a workbook with 2 spreadsheets
            var excel = new Application();
            excel.DisplayAlerts = false;
            Workbook storyWb = excel.Workbooks.Add(1);
            Worksheet itemSheet = (Worksheet)storyWb.Sheets[1];
            var relationshipSheet = storyWb.Worksheets.Add(Type.Missing, storyWb.Worksheets[storyWb.Worksheets.Count], 1, XlSheetType.xlWorksheet) as Worksheet;
            relationshipSheet.Name = "relationShipSheet";
            // Insert data into 2 spreadsheets
            RelationshipSheet(story, relationshipSheet);
            ItemSheet(story, itemSheet);

            // Convert the 2 csv files into a xlsx file
            string[] paths = new string[2] {fileLocation + "\\itemFile.csv" ,
                fileLocation + "\\relationshipFile.csv"};
            MergeWorkbooks(fileLocation + "\\combine.xlsx", paths);

            storyWb.Close();
            excel.Quit();
        }
        // Grabs all relationship data between 2 items.
        private static void RelationshipSheet(Story Story, Worksheet relationshipSheet)
        {
            // file path variable
            var fileLocation = System.IO.Directory.GetParent
                (System.IO.Directory.GetParent(Environment.CurrentDirectory)
                .ToString()).ToString();

            // Header Line
            relationshipSheet.Cells[1, "A"].Value2 = "Item 1";
            relationshipSheet.Cells[1, "B"].Value2 = "Item 2";
            relationshipSheet.Cells[1, "C"].Value2 = "Direction";
            var count = 2;

            // Parse through relationship data
            foreach (var line in Story.Relationships)
            {
                relationshipSheet.Cells[count, "A"] = line.Item1.Name;
                relationshipSheet.Cells[count, "B"] = line.Item2.Name;
                relationshipSheet.Cells[count, "C"] = line.Direction.ToString();
                count++;
            }
            //Write data to file
            relationshipSheet.SaveAs(fileLocation + "\\relationshipFile.csv", XlFileFormat.xlCSVWindows);
            Console.WriteLine("Relationship file written");

        }
        //Grabs all item data with their attributes.
        private static void ItemSheet(Story Story, Worksheet itemSheet)
        {
            // file location for output
            var fileLocation = System.IO.Directory.GetParent
                (System.IO.Directory.GetParent(Environment.CurrentDirectory).ToString()).ToString();
            // Initial Header list
            var headList = new List<string> { "Name", "Description", "Category", "Start", "Duration", "Resources", "Tags", "Panels", "Subcategory", "AttCount", "cat_color", "file_path" };
            // Grabs the attributes of the story
            var attData = Story.Attributes;
            // Grabs the categories of the story
            var catData = Story.Categories;
            // Filters the default attributes from the story
            var attList = new List<SC.API.ComInterop.Models.Attribute>();
            Regex regex = new Regex(@"none|None|Sample");
            var attCount = 0;
            foreach (var att in attData)
            {
                // Checks to see if attribute header is a default attritube.
                Match match = regex.Match(att.Name);
                if (!match.Success)
                {
                    // Adds non-default attribute to the List and to the header line
                    attList.Add(att);
                    attCount++;
                    headList.Add(att.Name + "|" + att.Type + "|" + att.Description);
                }
            }
            //Inserts headlist to first row of the sheet
            string[] header = headList.ToArray();
            char l = (char)((65) + (header.Length - 1));
            char n = 'o';
            // If column length is greater than 26, Add a "A" before every letter
            if (header.Length > 26)
            {
                n = (char)((65) + (header.Length - 26 - 1));
            }
            if (header.Length > 26)
            {
                itemSheet.Range[itemSheet.Cells[1, "A"], itemSheet.Cells[1, l.ToString() + n.ToString()]].Value2 = header;
            }
            else
            {
                itemSheet.Range[itemSheet.Cells[1, "A"], itemSheet.Cells[1, l.ToString()]].Value2 = header;
            }
            var itemCount = 2;
            // Goes through items in category order
            foreach (var cat in catData)
            {
                foreach (var item in Story.Items)
                {
                    var hasFile = false;
                    // check to see if category matches item category
                    if (item.Category.Name == cat.Name)
                    {
                        // Creates the initial list for the item 
                        var itemList = new List<string> { item.Name, item.Description, item.Category.Name, item.StartDate.ToString(), item.DurationInDays.ToString() };

                        //Goes through the item's resources
                        var resLine = "";
                        if (item.Resources.Length > 0)
                        {
                            foreach (var res in item.Resources)
                            {
                                //downloads resource file if there is a file extension to file
                                if (res.FileExtension != null)
                                {
                                    res.DownloadFile(fileLocation + "\\Files\\" + res.Name + res.FileExtension);
                                    resLine += res.Name + "~" + res.Name + "*" + res.FileExtension + "|";
                                    hasFile = true;
                                }
                                // Gets the url for a website
                                else
                                {
                                    resLine += res.Name + "~" + res.Url + "|";
                                }

                            }
                        }
                        // Item has no resources
                        else
                        {
                            resLine = "null";
                        }
                        // Add the resource data to list
                        itemList.Add(resLine);
                        // Adds the tags to the list
                        var tagLine = "";
                        if (item.Tags.Length > 0)
                        {
                            foreach (var tag in item.Tags)
                            {
                                tagLine += tag.Text + "|";
                            }
                        }
                        else
                        {
                            tagLine = "null";
                        }

                        itemList.Add(tagLine);
                        // Adds the panels to the list
                        var panLine = "";
                        // Check to see if item has any panels containing data
                        var dataCount = 0;
                        foreach (var pan in item.Panels)
                        {
                            // Check to see if panel data is empty
                            if (pan.Data.ToString() != "_EMPTY_")
                            {
                                dataCount++;
                                panLine += pan.Title + "@" + pan.Type + "@" + pan.Data + "|";
                            }

                        }
                        if (dataCount == 0)
                        {
                            panLine = "null";
                        }
                        itemList.Add(panLine);
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
                        itemList.Add(subLine);
                        itemList.Add(attCount.ToString());
                        //Adds the color of the category to the list
                        var colors = (cat.Color.A + "|" + cat.Color.R + "|" + cat.Color.G + "|" + cat.Color.B).ToString();
                        itemList.Add(colors);
                        // check to see if item has a image based off the sharpcloud image url
                        Regex zeroImage = new Regex(@"00000000");
                        Match zeroMatch = zeroImage.Match(item.ImageUri.ToString());
                         // Downloads image to folder if url is not all 0s
                        if (!zeroMatch.Success)
                        {
                            using (WebClient client = new WebClient())
                            {
                                client.DownloadFile(item.ImageUri, (fileLocation +
                                    "\\" + "Files" + "\\" + item.Name + ".jpg"));
                                itemList.Add(fileLocation + "\\" + "Files" + "\\");
                            }
                            hasFile = true;
                        }
                        if(hasFile == true)
                        {
                            itemList.Add(fileLocation + "\\" + "Files" + "\\");
                        }
                        else
                        {
                            itemList.add("null");   
                        }
                        // Adds the attributes to the item
                        foreach (var att in attList)
                        {
                            switch (att.Type.ToString())
                            {
                                case "Text":
                                    itemList.Add(item.GetAttributeValueAsText(att));
                                    break;
                                case "Numeric":
                                    itemList.Add(item.GetAttributeValueAsDouble(att).ToString());
                                    break;
                                case "Date":
                                    itemList.Add(item.GetAttributeValueAsDate(att).ToString());
                                    break;
                                case "List":
                                    itemList.Add(item.GetAttributeValueAsText(att));
                                    break;
                                case "Location":
                                    itemList.Add(item.GetAttributeValueAsText(att));
                                    break;
                            }
                        }
                        
                        // Adds entire list to the row for the item.
                        string[] itemLine = itemList.ToArray();
                        // If item column is greater than 26, Add a "A" before all letters
                        if (header.Length > 26)
                        {
                            itemSheet.Range[itemSheet.Cells[itemCount, "A"], itemSheet.Cells[itemCount, l.ToString() + n.ToString()]].Value2 = itemLine;
                        }
                        else
                        {
                            itemSheet.Range[itemSheet.Cells[itemCount, "A"], itemSheet.Cells[itemCount, l.ToString()]].Value2 = itemLine;
                        }
                        itemCount++;
                    }

                }
            }

            // Writes file to disk
            itemSheet.SaveAs(fileLocation + "\\itemFile.csv", XlFileFormat.xlCSVWindows);
            Console.WriteLine("ItemFile Written");
        }

        // method created by HuBeZa https://stackoverflow.com/a/32310557
        private static void MergeWorkbooks(string destinationFilePath, params string[] sourceFilePaths)
        {
            // Create a new workbook (index=1) and open source workbooks (index=2,3,...)
            var app = new Application();
            app.DisplayAlerts = false;
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
