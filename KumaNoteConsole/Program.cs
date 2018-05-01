using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;

namespace KumaNoteConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            //// skipping error checking, just demonstrating using these APIs
            //var app = new Application();

            //// get the hierarchy
            //string xmlHierarchy;
            //app.GetHierarchy(null, HierarchyScope.hsPages, out xmlHierarchy);

            //Console.WriteLine("Hierarchy:\n" + xmlHierarchy);

            //// now find the current page, print out its ID
            //var xdoc = XDocument.Parse(xmlHierarchy);
            //var ns = xdoc.Root.Name.Namespace;

            //var pageId = app.Windows.CurrentWindow.CurrentPageId;
            //Console.WriteLine("Current Page ID: " + pageId);

            //// get the page content, print it out
            //string xmlPage;
            //app.GetPageContent(pageId, out xmlPage);
            //Console.WriteLine("Page XML:\n" + xmlPage);

            //// sample - this is how to update content - normally you would modify the xml.
            //app.UpdatePageContent(xmlPage);

            //// bonus - if there are any images, get the binary content of the first one
            //var xPage = XDocument.Parse(xmlPage);
            //var xImage = xPage.Descendants(ns + "Image").FirstOrDefault();
            //if (xImage != null)
            //{
            //    var xImageCallbackID = xImage.Elements(ns + "CallbackID").First();
            //    var imageId = xImageCallbackID.Attribute("callbackID").Value;
            //    string base64Out;
            //    app.GetBinaryPageContent(pageId, imageId, out base64Out);

            //    Console.WriteLine("Image found, base64 data is:\n" + base64Out);
            //}

            var onenoteApp = new Application();

            /*-------------------------------notebook---------------------------------*/
            Console.WriteLine("---------------------------------------NOTE START---------------------------------");

            var currentNotebookId = onenoteApp.Windows.CurrentWindow.CurrentNotebookId;    
            Console.WriteLine("CurrentNotebookId:" + currentNotebookId);

            string xmlHierarchy;
            onenoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xmlHierarchy);
            Console.WriteLine("xmlHierarchy:" + xmlHierarchy);

            var xdoc = XDocument.Parse(xmlHierarchy);
            var ns = xdoc.Root.Name.Namespace;
            Console.WriteLine(ns);

            var notebook = xdoc.Descendants(ns + "Notebook").FirstOrDefault();
            if (notebook != null) {
                Console.WriteLine(notebook);
            }
            Console.WriteLine("------------------------------------------------------");
            foreach (XElement note in xdoc.Descendants(ns + "Notebook")) {
                Console.WriteLine(note);
            }

            Console.WriteLine("------------------------------------------------------");



            Console.WriteLine("---------------------------------------NOTE END---------------------------------");

            /*-------------------------------page---------------------------------*/
            Console.WriteLine("---------------------------------------PAGE START---------------------------------");

            var currentPageId = onenoteApp.Windows.CurrentWindow.CurrentPageId;
            Console.WriteLine("CurrentPageId:" + currentPageId);

            onenoteApp.GetHierarchy(currentNotebookId, HierarchyScope.hsPages, out xmlHierarchy);
            Console.WriteLine("xmlHierarchy:" + xmlHierarchy);

            xdoc = XDocument.Parse(xmlHierarchy);
            ns = xdoc.Root.Name.Namespace;
            Console.WriteLine(ns);

            //var page = xdoc.Descendants(ns + "Page").FirstOrDefault();
            //if (page != null)
            //{
            //    Console.WriteLine(page);
            //}
            Console.WriteLine("------------------------------------------------------");
            foreach (XElement page in xdoc.Descendants(ns + "Page"))
            {
                Console.WriteLine(page);
            }

            Console.WriteLine("------------------------------------------------------");

            xmlHierarchy = "<?xml version=\"1.0\"?>\r\n<one:Notebook xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" name=\"TestEnviroment\" nickname=\"TestEnviroment\" ID=\"{7D8724A3-C597-4F21-8B22-D8F1DBE6E0C6}{1}{B0}\" path=\"https://d.docs.live.net/e6668350c8c57e54/01.private/03.笔记/TestEnviroment/\" lastModifiedTime=\"2018-04-30T16:16:25.000Z\" color=\"#B49EDE\" isCurrentlyViewed=\"true\">\r\n    <one:Section name=\"TestSection1\" ID=\"{CA02267A-8F1C-49A2-A6A2-D109D4E25D65}{1}{B0}\" path=\"https://d.docs.live.net/e6668350c8c57e54/01.private/03.笔记/TestEnviroment/TestSection1.one\" lastModifiedTime=\"2018-04-30T16:16:25.000Z\" color=\"#8AA8E4\" isCurrentlyViewed=\"true\">\r\n        <one:Page ID=\"{CA02267A-8F1C-49A2-A6A2-D109D4E25D65}{1}{E1948215333177470140691997556298680726132841}\" name=\"TestPage1.3\" dateTime=\"2018-04-30T16:16:08.000Z\" lastModifiedTime=\"2018-04-30T16:16:19.000Z\" pageLevel=\"1\" isCurrentlyViewed=\"true\" />\r\n        <one:Page ID=\"{CA02267A-8F1C-49A2-A6A2-D109D4E25D65}{1}{E1951027410831158810631971022623693182670241}\" name=\"TestPage1.2\" dateTime=\"2018-04-30T16:16:07.000Z\" lastModifiedTime=\"2018-04-30T16:16:11.000Z\" pageLevel=\"1\" />\r\n        <one:Page ID=\"{CA02267A-8F1C-49A2-A6A2-D109D4E25D65}{1}{E19552480321094610014420171768064772873089501}\" name=\"TestPage1.1\" dateTime=\"2018-04-30T15:07:36.000Z\" lastModifiedTime=\"2018-04-30T15:09:42.000Z\" pageLevel=\"1\" />\r\n    </one:Section>\r\n    <one:Section name=\"TestSection2\" ID=\"{F9B03AC0-B030-4E81-9DAE-D51AB80B3AF3}{1}{B0}\" path=\"https://d.docs.live.net/e6668350c8c57e54/01.private/03.笔记/TestEnviroment/TestSection2.one\" lastModifiedTime=\"2018-04-30T15:09:26.000Z\" color=\"#91BAAE\">\r\n        <one:Page ID=\"{F9B03AC0-B030-4E81-9DAE-D51AB80B3AF3}{1}{E19529780536953855570320178099339487768224521}\" name=\"Page2.1\" dateTime=\"2018-04-30T15:09:14.000Z\" lastModifiedTime=\"2018-04-30T15:09:21.000Z\" pageLevel=\"1\" />\r\n    </one:Section>\r\n</one:Notebook>";

            onenoteApp.UpdateHierarchy(xmlHierarchy);


            Console.WriteLine("---------------------------------------PAGE END---------------------------------");


            /*-------------------------------section group---------------------------------*/
            var currentSectionGroupId = onenoteApp.Windows.CurrentWindow.CurrentSectionGroupId;
            Console.WriteLine("CurrentSectionGroupId:" + currentSectionGroupId);

            /*-------------------------------section---------------------------------*/

            var currentSectionId = onenoteApp.Windows.CurrentWindow.CurrentSectionId;
            Console.WriteLine("CurrentSectionId:" + currentSectionId);


        }
    }
}
