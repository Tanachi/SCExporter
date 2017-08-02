using System;
using System.IO;
using System.Configuration;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Net;
using System.Threading;
using OfficeOpenXml;
using System.Linq;
// Grabs story info from Sharpcloud and converts the data into a relationship sheet and item sheet.
namespace SCExporter
{
    class Program
    {
        public class Sharp
        {

            public SharpCloudApi sc { get; set; }
            public Story story { get; set; }
            public ExcelWorksheet sheet { get; set; }
            public int order { get; set; }
            public int attCount { get; set; }
            public List<SC.API.ComInterop.Models.Attribute> attList {get; set;}
            public int sheetLine { get; set; }
            
        }
        static Thread firstHalf,downloader;
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
            //Create EPPlus and create a workbook with 2 spreadsheets
            FileInfo newFile = new FileInfo(fileLocation + "\\combine.xlsx");
            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheets
            var itemSheet = pck.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Items");
            if (itemSheet == null)
            {
                itemSheet = pck.Workbook.Worksheets.Add("Items");
            }
            var relationshipSheet = pck.Workbook.Worksheets.First();
            if (pck.Workbook.Worksheets.Count > 1)
                relationshipSheet = pck.Workbook.Worksheets.ElementAt(1);
            if (relationshipSheet == itemSheet)
            {
                relationshipSheet = pck.Workbook.Worksheets.Add("RelationshipSheet");
            }
            //Setting up the threads
            object a, b;
            Sharp first = new Sharp();
            Sharp second = new Sharp();
            first.sc = sc; first.story = story; first.sheet = itemSheet;
            first.order = 0;
            second.story = story;
            second.order = 1;
            a = first;
            b = second;
            // Initial Item Header list
            var headList = new List<string> { "Name", "Description", "Category", "Start", "Duration",
                "Resources", "Tags", "Panels", "Subcategory", "AttCount", "cat_color", "file_path" };
            // Filters the default attributes from the story
            var attData = story.Attributes;
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
            var go = 1;
            foreach (var head in headList)
            {
                itemSheet.Cells[1, go].Value = head;
                go++;
            }
            first.attList = attList;
            first.sheetLine = 2;
            first.attCount = attCount;

            // Insert data into 2 spreadsheets
            firstHalf = new Thread(new ParameterizedThreadStart(downloadFiles));
            downloader = new Thread(new ParameterizedThreadStart(downloadFiles));

            firstHalf.Start(a);
            downloader.Start(b);
            ItemSheet(a);
            RelationshipSheet(story, relationshipSheet);

            // Save the workbook
            pck.SaveAs(newFile);
        }
        // Grabs all relationship data between 2 items
        private static void RelationshipSheet(Story Story, OfficeOpenXml.ExcelWorksheet relationshipSheet)
        {
            // file path variable
            // Header Line
            relationshipSheet.Cells["A1"].Value = "Item 1";
            relationshipSheet.Cells["B1"].Value = "Item 2";
            relationshipSheet.Cells["C1"].Value = "Direction";
            var count = 2;
            // Parse through relationship data
            foreach (var line in Story.Relationships)
            {
                var go = 1;
                relationshipSheet.Cells[count, go].Value = line.Item1.Name;
                go++;
                relationshipSheet.Cells[count, go].Value = line.Item2.Name;
                go++;
                relationshipSheet.Cells[count, go].Value = line.Direction.ToString();
                count++;
            }
            //Write data to file
            Console.WriteLine("Relationship sheet written");
        }
        //Grabs all item data with their attributes
        private static void ItemSheet(object Sharp)
        {
            
            var fileLocation = System.IO.Directory.GetParent
                (System.IO.Directory.GetParent(Environment.CurrentDirectory)
                .ToString()).ToString();
            Sharp sharp = Sharp as Sharp;
            var story = sharp.story;
            var sc = sharp.sc;
            var catData = story.Categories;
            var attList = sharp.attList;
            var attCount = sharp.attCount;
            var itemSheet = sharp.sheet;
            var order = sharp.order;
            var sheetLine = sharp.sheetLine;
            // Goes through items in category order
            foreach (var cat in catData)
            {      
                foreach(var item in story.Items)
                {
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
                                    resLine += res.Name + "~" + res.Name + "*" + res.FileExtension + "|";
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
                        itemList.Add(fileLocation + "\\" + "Files" + "\\");
                        // Adds the attributes to the item
                        string[] itemLine = itemList.ToArray();
                        var go = 1;
                        foreach (var itemCell in itemLine)
                        {
                            itemSheet.Cells[sheetLine, go].Value = itemCell;
                            if (go == 5)
                            {
                                itemSheet.Cells[sheetLine, go].Value = Double.Parse(itemCell);
                                itemSheet.Cells[sheetLine, go].Style.Numberformat.Format = "#";
                            }
                            if (go == 10)
                            {
                                itemSheet.Cells[sheetLine, go].Value = Double.Parse(itemCell);
                            }
                                
                            go++;
                        }
                        // Adds the attributes to the item
                        foreach (var att in attList)
                        {
                            switch (att.Type.ToString())
                            {
                                case "Text":
                                    itemSheet.Cells[sheetLine, go].Value = item.GetAttributeValueAsText(att);
                                    break;
                                case "Numeric":
                                    itemSheet.Cells[sheetLine, go].Value = item.GetAttributeValueAsDouble(att);
                                    itemSheet.Cells[sheetLine, go].Style.Numberformat.Format = "0.00";
                                    break;
                                case "Date":
                                    itemSheet.Cells[sheetLine, go].Value = item.GetAttributeValueAsDate(att);
                                    break;
                                case "List":
                                    itemSheet.Cells[sheetLine, go].Value = item.GetAttributeValueAsText(att);
                                    break;
                                case "Location":
                                    itemSheet.Cells[sheetLine, go].Value = item.GetAttributeValueAsText(att);
                                    break;
                            }
                            go++;
                        }
                        sheetLine++;
                    }
                }
            }

            // Writes file to disk
            Console.WriteLine("ItemSheet Written");
        }
        private static void downloadFiles(object Sharp)
        {
            var fileLocation = System.IO.Directory.GetParent
                (System.IO.Directory.GetParent(Environment.CurrentDirectory)
                .ToString()).ToString();
            Sharp sharp = Sharp as Sharp;
            var story = sharp.story;
            int order = sharp.order;
            var a = 0;
            var b = (story.Items.Length / 2) - 1;
            if(order > 0)
            {
                a = story.Items.Length / 2;
                b = story.Items.Length;
            }
            for (var i = a; i < b;i++)
            {
                //Goes through the item's resources
                if (story.Items[i].Resources.Length > 0)
                {
                    foreach (var res in story.Items[i].Resources)
                    {
                        //downloads resource file if there is a file extension to file
                        if (res.FileExtension != null)
                        {
                            res.DownloadFile(fileLocation + "\\Files\\" + res.Name + res.FileExtension);
                        }
                    }
                }
                // check to see if item has a image based off the sharpcloud image url
                Regex zeroImage = new Regex(@"00000000");
                Match zeroMatch = zeroImage.Match(story.Items[i].ImageUri.ToString());
                // Downloads image to folder if url is not all 0s
                if (!zeroMatch.Success)
                {

                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile(story.Items[i].ImageUri, (fileLocation + "\\" + "Files" + "\\" + story.Items[i].Name + ".jpg"));
                    }
                }
            }
            Console.WriteLine("Files Downloaded");
        }
        
    }
}
