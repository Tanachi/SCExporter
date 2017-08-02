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
using OfficeOpenXml;
using System.Linq;
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
            var templateid = ConfigurationManager.AppSettings["templateid"];
            var userid = ConfigurationManager.AppSettings["user"];
            var passwd = ConfigurationManager.AppSettings["pass"];
            var URL = ConfigurationManager.AppSettings["URL"];
            var storyID = ConfigurationManager.AppSettings["story"];

            // Login and get story data from Sharpcloud
            var sc = new SharpCloudApi(userid, passwd, URL);
            var story = sc.LoadStory(storyID);
            //Create excel application and create a workbook with 2 spreadsheets
            FileInfo newFile = new FileInfo(fileLocation + "\\combine.xlsx");

            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheet
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
            // Insert data into 2 spreadsheets
            RelationshipSheet(story, relationshipSheet);
            ItemSheet(story, itemSheet);

            pck.SaveAs(newFile);
        }
        // Grabs all relationship data between 2 items.
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
        //Grabs all item data with their attributes.

        
        private static void ItemSheet(Story Story, OfficeOpenXml.ExcelWorksheet itemSheet)
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
            var go = 1;
            foreach (var head in headList)
            {

                itemSheet.Cells[1, go].Value = head;
                go++;
            }
            //Inserts headlist to first row of the sheet
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
                                client.DownloadFile(item.ImageUri, (fileLocation + "\\" + "Files" + "\\" + item.Name + ".jpg"));
                            }
                            hasFile = true;
                        }
                        if (hasFile)
                        {
                            itemList.Add(fileLocation + "\\" + "Files" + "\\");
                        }
                        else
                        {
                            itemList.Add("null");
                        }
                        string[] itemLine = itemList.ToArray();
                        var go = 1;
                        foreach (var itemCell in itemLine)
                        {
                            itemSheet.Cells[itemCount, go].Value = itemCell;
                            if (go == 5)
                            {
                                itemSheet.Cells[itemCount, go].Value = Double.Parse(itemCell);
                                itemSheet.Cells[itemCount, go].Style.Numberformat.Format = "#";
                            }
                            if (go == 10)
                            {
                                itemSheet.Cells[itemCount, go].Value = Double.Parse(itemCell);
                            }
                                
                            go++;
                        }
                        // Adds the attributes to the item
                        foreach (var att in attList)
                        {
                            switch (att.Type.ToString())
                            {
                                case "Text":
                                    itemSheet.Cells[itemCount, go].Value = item.GetAttributeValueAsText(att);
                                    break;
                                case "Numeric":
                                    itemSheet.Cells[itemCount, go].Value = item.GetAttributeValueAsDouble(att);
                                    itemSheet.Cells[itemCount, go].Style.Numberformat.Format = "0.00";
                                    break;
                                case "Date":
                                    itemSheet.Cells[itemCount, go].Value = item.GetAttributeValueAsDate(att);
                                    break;
                                case "List":
                                    itemSheet.Cells[itemCount, go].Value = item.GetAttributeValueAsText(att);
                                    break;
                                case "Location":
                                    itemSheet.Cells[itemCount, go].Value = item.GetAttributeValueAsText(att);
                                    break;
                            }
                            go++;
                        }
                        itemCount++;
                    }

                }
            }

            // Writes file to disk
            Console.WriteLine("ItemSheet Written");
        }
        
    }
}
