using System;
using System.IO;
using System.Configuration;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;

// Grabs story info from Sharpcloud and converts the data into a relationship sheet and item sheet.
namespace SCExporter
{
    class Program
    {
        static void Main(string[] args)
        {
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

            // Sort data into 2 different spreadsheets and writes them to disk.
            RelationshipSheet(story);
            ItemSheet(story);
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
            foreach (var att in attData)
            {
                // Checks to see if attribute header is a default attritube.
                Match match = regex.Match(att.Name);
                if (!match.Success)
                {
                    // Adds non-default attribute to the List and to the header line
                    attList.Add(att);
                    itemFile += "," + att.Name;
                }
            }
            // End of Initial Header line
            itemFile += "," + "Tags" + Environment.NewLine;
            // Goes through array of the story's item
            foreach (var cat in catData)
            {
                foreach (var item in Story.Items)
                {
                    if (item.Category.Name == cat.Name)
                    {
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
                            tagLine += tag.Text;
                        }
                        itemLine += "," + tagLine + Environment.NewLine;
                        // Adds the line to the file
                        itemFile += itemLine;
                    }
                }
            }
            // Writes file to disk
            File.WriteAllText(System.IO.Directory.GetParent(System.IO.Directory.GetParent(Environment.CurrentDirectory).ToString()).ToString() + "\\itemFile.csv", itemFile);
            Console.WriteLine("ItemFile Written");
        }
    }
}
