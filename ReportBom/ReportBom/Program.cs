using Newtonsoft.Json;
using SolidEdgeCommunity;
using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;



namespace ReportBom
{


    class Program
    {
        

        [STAThread]
        static void Main()
        {
            SolidEdgeFramework.Application application = null;


            try
            {
                OleMessageFilter.Register();

                // Connect to a running instance of Solid Edge.
                 application = SolidEdgeUtils.Connect();

                // Connect to the active assembly document.
                var assemblyDocument = application.GetActiveDocument<SolidEdgeAssembly.AssemblyDocument>(false);

                // Optional settings you may tweak for performance improvements. Results may vary.
                application.DelayCompute = true;
                application.DisplayAlerts = false;
                application.Interactive = false;
                application.ScreenUpdating = false;

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                    return;
                }
                xlApp.Visible = true;

                // Open workbook
                Workbook wb = xlApp.Workbooks.Open("Y:\\Common\\Engineering\\001_Vraagbaak\\gijs\\BOM.xlsx");

                // Open worksheet
                Worksheet ws = (Worksheet)wb.Worksheets[2];

                if (ws == null)
                {
                    Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
                }

                int i = 2;
                string sRow = "";
                object[] oBomValues = new Object[1];

                sRow = i.ToString();

                oBomValues[0] = "Keywords";
                Range aRange = ws.get_Range("C" + sRow);
                aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, oBomValues);

                oBomValues[0] = "Document Number";
                aRange = ws.get_Range("D" + sRow);
                aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, oBomValues);

                oBomValues[0] = "Title";
                aRange = ws.get_Range("E" + sRow);
                aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, oBomValues);

                oBomValues[0] = "Quantity";
                aRange = ws.get_Range("F" + sRow);
                aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, oBomValues);

                i++;

                if (assemblyDocument != null)
                {
                    var rootBomItem = new BomItem();
                    rootBomItem.FileName = assemblyDocument.FullName;

                    // Write Name of rootBomItem to excel
                    oBomValues[0] = assemblyDocument.DisplayName;
                    aRange = ws.get_Range("A1");
                    aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, oBomValues);

                    // Begin the recurisve extraction process.
                    PopulateBom(0, assemblyDocument, rootBomItem);

                    // Write each BomItem to console.
                    foreach (var bomItem in rootBomItem.AllChildren)
                    {

                        Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}", bomItem.Level, bomItem.DocumentNumber, bomItem.Revision, bomItem.Title, bomItem.Quantity);
                        sRow = i.ToString();

                        oBomValues[0] =  bomItem.Level;
                        aRange = ws.get_Range("C" + sRow);
                        aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, oBomValues);

                        oBomValues[0] = bomItem.DocumentNumber;
                        aRange = ws.get_Range("D" + sRow);
                        aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, oBomValues);

                        oBomValues[0] = bomItem.Title;
                        aRange = ws.get_Range("E" + sRow);
                        aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, oBomValues);

                        oBomValues[0] = bomItem.Quantity;
                        aRange = ws.get_Range("F" + sRow);
                        aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, oBomValues);

                        i++;

                    }

                    // Demonstration of how to save the BOM to various formats.

                    // Define the Json serializer settings.
                    var jsonSettings = new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    };

                    // Convert the BOM to JSON.
                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(rootBomItem, Newtonsoft.Json.Formatting.Indented, jsonSettings);

                    wb.RefreshAll();
                    wb.SaveAs(assemblyDocument.Path + "\\_BOM.xlsx");

                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (application != null)
                {
                    application.DelayCompute = false;
                    application.DisplayAlerts = true;
                    application.Interactive = true;
                    application.ScreenUpdating = true;
                }

                OleMessageFilter.Unregister();
            }
        }

   

            static void PopulateBom(int level, SolidEdgeAssembly.AssemblyDocument assemblyDocument, BomItem parentBomItem)
        {
            // Increment level (depth).
            level++;

            // This sample BOM is not exploded. Define a dictionary to store unique occurrences.
            //Dictionary<string, SolidEdgeAssembly.Occurrence> uniqueOccurrences = new Dictionary<string, SolidEdgeAssembly.Occurrence>();
            


            // Loop through the unique occurrences.
            foreach (SolidEdgeAssembly.Occurrence occurrence in assemblyDocument.Occurrences)  // uniqueOccurrences.Values.ToArray())
            {
                // Filter out certain occurrences.
                if (!occurrence.IncludeInBom) { continue; }
                if (occurrence.IsPatternItem) { continue; }
                if (occurrence.OccurrenceDocument == null) { continue; }

                // Create an instance of the child BomItem.
                var bomItem = new BomItem(occurrence, level);

                // Add the child BomItem to the parent.
                parentBomItem.Children.Add(bomItem);

                if (bomItem.IsSubassembly == true)
                {
                    // Sub Assembly. Recurisve call to drill down.
                    PopulateBom(level, (SolidEdgeAssembly.AssemblyDocument)occurrence.OccurrenceDocument, bomItem);
                }
            }
        }
    }
}
