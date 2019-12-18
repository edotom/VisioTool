using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Visio;
using Nice.VisioTool.Structs;
using System.IO;


namespace Demo
{

    class Program
    {
        // We need to the scope of these fields to be global so that all of the
        // methods in the class can use them.
        private Document _vsoStencil;
        private Page _vsoPage;
        private Master _vsoMaster;

        private List<AbstractShape> shapes = new List<AbstractShape>();
        private Dictionary<int, string> shapeList = null;

        static void Main(string[] args)
        {
            new Program()._Main();

            // Stop program execution during debugging.
            Console.ReadKey();
        }

        private void _Main()
        {
            Application vsoApp = null;
            Document vsoDoc = null;
            StringBuilder contentsForFile = new StringBuilder();
            string visioDocument = @"c:\temp\test.vsdx";
            string folderPath = "c:\\temp";
            string fileName = "test.txt";

            try
            {
                vsoApp = new Application();
                int pageNumber = 1;


                vsoDoc = vsoApp.Documents.Open(visioDocument);

                //_vsoPage = vsoDoc.Pages[pageNumber];
                GetAllPages(vsoDoc);
                // Add the start shape to the shapeList so that we don't crawl it again.
                shapeList = new Dictionary<int, string>();


                // Make the first call to GetConnectedShapes.
                //GetConnectedShapes(_vsoPage.Shapes[1]);
                // We're done with the document so we can save it.
                foreach (var shape in shapes)
                {
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    contentsForFile.Append($"Page: {shape.PageName}, Type: Page"+ Environment.NewLine);
                    Console.WriteLine($"Page: {shape.PageName}, Type: Page");
                    contentsForFile.Append($"Shape: {shape.Name}, Type: {shape.Type}" + Environment.NewLine);
                    Console.WriteLine($"Shape: {shape.Name}, Type: {shape.Type}");
                    foreach (var connector in shape.ConnectedTo)
                    {
                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        contentsForFile.Append($"    -> Connected to: {connector}" + Environment.NewLine);
                        Console.WriteLine($"    -> Connected to: {connector}");
                    }

                }

                Console.WriteLine($"Visio file processed.");
                Console.WriteLine($"Select an option.");
                Console.WriteLine($"1. Export to text file");
                Console.WriteLine($"2. Export to excel file");
                Console.WriteLine($"3. End");

                string option = Console.ReadLine();
                Nice.VisioTool.Tools.TextTools oTool = new Nice.VisioTool.Tools.TextTools(folderPath, fileName);
                Console.WriteLine(string.Format("selected option {0}", option));               
                switch (option)
                {
                    case "1":
                        if (!File.Exists(string.Format("{0}\\{1}", folderPath, fileName)))
                        {
                            oTool.CreateFile();
                        }
                        oTool.WriteToFile(contentsForFile.ToString());
                        break;
                    case "2":

                        break;
                    default:
                        Environment.Exit(1);
                        break;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Clean up now that we're done with everything.
                if (vsoDoc != null)
                {
                    vsoDoc.Close();
                }
                if (_vsoStencil != null)
                {
                    _vsoStencil.Close();
                }
                if (vsoApp != null)
                {
                    vsoApp.Quit();
                }
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void GetConnectedShapes(Shape vsoShape, Page vsoPage)
        {
            Shape vsoOutShape;
            int outNodes;
            bool beenChecked = false;
            try
            {
                AbstractShape absShape = new AbstractShape();
                absShape.PageName = vsoPage.Name;
                absShape.Name = vsoShape.Text;
                absShape.Type = vsoShape.NameU;
                absShape.UniqueID = vsoShape.ID;

                // Get an array of all the shapes connected to the shape passed into the argument.
                var shapeIDArray =
                    vsoShape.ConnectedShapes(
                       VisConnectedShapesFlags.visConnectedShapesOutgoingNodes,
                        "");

                if (shapeIDArray.Length > 0)
                {
                    outNodes = shapeIDArray.Length;
                    for (var node = 0; node < outNodes; node++)
                    {

                        //vsoOutShape = _vsoPage.Shapes.get_ItemFromID((int)shapeIDArray.GetValue(node));
                        vsoOutShape = vsoPage.Shapes.get_ItemFromID((int)shapeIDArray.GetValue(node));
                        absShape.ConnectedTo.Add(vsoOutShape.Text);
                        shapes.Add(absShape);

                        // We need to see if the connected shape has shapes connected to it,
                        // so we'll get array of all the shapes connected to it.
                        var outShapeIDArray = vsoOutShape.ConnectedShapes(
                            VisConnectedShapesFlags.visConnectedShapesOutgoingNodes,
                            "");

                        string value;
                        if (shapeList != null && shapeList.TryGetValue(vsoOutShape.ID, out value))
                        {
                            beenChecked = true;
                        };

                        // The connected shape has shapes attached to it and it hasn't been,
                        // checked yet, so we'll crawl it next.
                        if ((outShapeIDArray.Length > 0) &
                        (!beenChecked && shapeList != null))
                        {
                            shapeList.Add(vsoOutShape.ID, vsoOutShape.Text);
                            GetConnectedShapes(vsoOutShape, vsoPage);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
            }

        }

        private void GetAllPages(Document vsoDoc)
        {
            foreach (Page InPage in vsoDoc.Pages)
            {

                new WalkingConnections().ShowPageConnections(InPage);
                //Console.WriteLine(string.Format("Page is:{0} and ID is : {1}", InPage.Name, InPage.ID));//cheched working
                foreach (Shape InShape in InPage.Shapes)
                {
                   // Console.WriteLine(string.Format("Shape is:{0} and ID is : {1}", InShape.Name, InShape.ID));
                    GetConnectedShapes(InShape, InPage);

                }
                //if(_vsoPage.Shapes.Count > 0)
                //foreach (Shape InShape in InPage.Shapes)
                //{

                //     //   if(InPage.Shapes.Count > 0)
                //    //    GetConnectedShapes(InShape, InPage);
                //    //    foreach (Connect InConnect in InShape.Connects)
                //   // {

                //   // }

                //    //Console.WriteLine(string.Format("Name is:{0} and ID is : {1} and ObjectType is: {2}", InShape.Name, InShape.ID, InShape.Type.ToString()));
                //    //foreach (long ID in InShape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesAllNodes, ""))
                //   // {
                //    //    Console.WriteLine(string.Format("ID is:{0}", ID.ToString()));
                //   // }
                //}


            }

        }

    }
}

