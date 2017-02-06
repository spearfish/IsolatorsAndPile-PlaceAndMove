#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace SOM.RevitTools.PlaceIsolator
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        //  Author:   Danny Bentley
        //  Date  :   9/16/16
        //  Objective : Place isolator elements using Excel XY coordinates

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            ExcelReader Excel = new ExcelReader();
            program Program = new program();
            // List for isolators 
            List<IsoObj> List_Isolators = new List<IsoObj>();
            // User Interface window
            Window_UI window = new Window_UI();

            // Add all levels to the UI to be picked. 
            window.LevelComboBox.ItemsSource = getLevels(doc);
            var res = window.ShowDialog();

            // Check if user clicked "OK" button, "Cancel" button or "Clear" button
            if (!(res.HasValue && res.Value)) return Result.Cancelled;
            // Selected level and file.
            string SelectedLevel = window.LevelComboBox.Text;
            string Filepath = window._Filepath;

            List_Isolators = Excel.Get_ExcelFileToList(Filepath);

            Dictionary<string, Element> dictionary = new Dictionary<string, Element>();
            // Get family elment and ID to a Dictionary. 
            try
            {
                dictionary = Program.Get_ElementAndParamValues(doc);

            }
            catch { }

            

            foreach (IsoObj Isolator in List_Isolators)
            {
                Program.porgramIsolator(uiapp, doc, uidoc, Isolator, SelectedLevel, dictionary);
            }

            return Result.Succeeded;
        }

        //*****************************getLevels()*****************************
        // Get levels to be added to the UI.
        public List<string> getLevels(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> levels = collector.OfClass(typeof(Level)).ToElements();
            List<string> List_levels = new List<string>();
            foreach (Level level in levels)
            {
                List_levels.Add(level.Name);
            }
            return List_levels;
        }
    }
}
