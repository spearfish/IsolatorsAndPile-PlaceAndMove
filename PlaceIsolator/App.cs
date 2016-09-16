using System;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using System.Reflection;

namespace SOM.RevitTools.PlaceIsolator
{
    class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            //Create a add to SOM Tools ribbon tab
            try
            {
                a.CreateRibbonTab("LACMA");
            }
            catch (Exception)
            {

            }

            //Create Ribbon Panel
            RibbonPanel structuralpanel = a.CreateRibbonPanel("LACMA", "LACMA Isolator");

            //Add button to the panel
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButton sheetButton = structuralpanel.AddItem(new PushButtonData
                ("LACMA ISO", "ISOLATOR", thisAssemblyPath, "SOM.RevitTools.PlaceIsolator.Command")) as PushButton;

            //Tooltip 
            sheetButton.ToolTip = "Place Base Isolator in LACMA project.";

            ////Set globel directory
            var globePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "LACMA ISO.png");

            //Large image 
            Uri uriImage = new Uri(globePath);
            BitmapImage largeimage = new BitmapImage(uriImage);
            sheetButton.LargeImage = largeimage;


            a.ApplicationClosing += a_ApplicationClosing;

            //Set Application to Idling
            a.Idling += a_Idling;

            return Result.Succeeded;
        }

        void a_Idling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {

        }

        void a_ApplicationClosing(object sender, Autodesk.Revit.UI.Events.ApplicationClosingEventArgs e)
        {
            throw new NotImplementedException();
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}