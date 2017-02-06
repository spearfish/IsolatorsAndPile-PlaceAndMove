using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using DBLibrary;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;

namespace SOM.RevitTools.PlaceIsolator
{
    class program
    {
        //*****************************porgramIsolator()*****************************
        public string porgramIsolator(UIApplication uiApp, Document doc, UIDocument uiDocument,
                                  IsoObj Isolator, string SelectedLevel, Dictionary<string, Element> dictionary)
        {
            try
            {
                // Family name and type being used. 
                string fsFamilyName = "BASE ISOLATOR - (I SF)";
                string fsFamilyType = "TYPE 1";
                // Selected level 
                string levelName = SelectedLevel;

                // Identification 
                string id = Isolator.ID;

                //Test to see if the dictionary contains the Element 
                if (dictionary.ContainsKey(id))
                {
                    // look up Key in dictionary to get Element. 
                    Element element = dictionary[id];
                    MoveInstance(doc, element, Isolator, SelectedLevel);
                }
                // If dictioary doesn't contian element it will create a new one. 
                else
                {
                    // Method to create family. 
                    Create_BaseIsolator(uiDocument, doc, fsFamilyName,
                                        fsFamilyType, levelName, Isolator);
                }
            }
            catch (Exception e)
            {
                string message = e.Message;
                return Result.Failed.ToString();
            }
            return Result.Succeeded.ToString();
        }

        //*****************************Get_ElementAndParamValues()*****************************
        public Dictionary<string, Element> Get_ElementAndParamValues(Document doc)
        {
            LibraryGetItems library = new LibraryGetItems();
            // Get Structural foundation elements 
            FilteredElementCollector dcs = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFoundation);
            //Build a Dictonary of elements and Keys of ID.
            int n = dcs.GetElementCount();
            Dictionary<string, Element> d = new Dictionary<string, Element>(n);

            foreach (Element dc in dcs)
            {
                // Get SOM ID parameter. 
                var SOMIDParam = dc.LookupParameter("SOM ID");
                string parameterVaule = library.GetParameterValue(SOMIDParam);

                d.Add(parameterVaule, dc);
            }
            return d;
        }

        //*****************************Create_BaseIsolator()*****************************
        public void Create_BaseIsolator(UIDocument uidoc, Document doc, string fsFamilyName,
                                        string fsFamilyType, string levelName, IsoObj Isolator)
        {
            // XY
            string xCoord = Isolator.X;
            string yCoord = Isolator.Y;
            // Rotation
            //string rotation = Isolator.R;
            //XY convert to double and divide by 12. 
            double x = double.Parse(xCoord) / 12;
            double y = double.Parse(yCoord) / 12;
            //double x = lib.MmToFoot(double.Parse(xCoord));
            //double y = lib.MmToFoot(double.Parse(yCoord));

            // Identification 
            string id = Isolator.ID;

            // LINQ to find the FamilySymbol by its type name.
            FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).
                                             OfClass(typeof(FamilySymbol)).
                                             Cast<FamilySymbol>()
                                         where (fs.Family.Name == fsFamilyName && fs.Name == fsFamilyType)
                                         select fs).First();
            // Get Level 
            Level level = GetLevel(doc, levelName);

            // Convert coordinates to double and create XYZ point.

            //double r = double.Parse(rotation);
            XYZ xyz = new XYZ(x, y, level.Elevation);
            XYZ xyz_A = new XYZ(x, y, level.Elevation + 2);
            Line axis = Line.CreateBound(xyz, xyz_A);


            // Create Isolator.
            using (Transaction t = new Transaction(doc, "Create Base Isolator"))
            {
                t.Start();
                if (!familySymbol.IsActive)
                {

                    // Ensure the family symbol is activated.
                    familySymbol.Activate();
                    doc.Regenerate();
                }

                // Create Base Isolator. 
                FamilyInstance IsolatorInstance = doc.Create.NewFamilyInstance(xyz, familySymbol, level, StructuralType.Footing);
                //FamilyInstance IsolatorInstance = doc.Create.NewFamilyInstance(xyz, familySymbol, hostedElement, level, StructuralType.Footing);

                // Rotate family elmement
                //IsolatorInstance.Location.Rotate(axis, r);

                // Parameter ID for tracking (MUST ADD PARAMETER TO PROJECT)
                Parameter SOMIDParam = IsolatorInstance.LookupParameter("SOM ID");
                SOMIDParam.Set(id).ToString();

                t.Commit();
            }
        }

        //*****************************MoveInstance()*****************************
        public void MoveInstance(Document doc, Element element, IsoObj Isolator, string levelName)
        {
            // get the column current location
            LocationPoint InstanceLocation = element.Location as LocationPoint;
            LibraryConvertUnits lib = new LibraryConvertUnits();

            XYZ oldPlace = InstanceLocation.Point;
            double Old_Rotation = InstanceLocation.Rotation;

            //Get Level
            Level level = GetLevel(doc, levelName);

            // XY
            string xCoord = Isolator.X;
            string yCoord = Isolator.Y;
            // Rotation
            string rotation = Isolator.R;

            double New_x = lib.MmToFoot(double.Parse(xCoord));
            double New_y = lib.MmToFoot(double.Parse(yCoord));
            double New_Rotation = double.Parse(rotation);

            XYZ New_xyz = new XYZ(New_x, New_y, level.Elevation);
            XYZ New_xyz_A = new XYZ(New_x, New_y, level.Elevation + 2);
            Line New_Axis = Line.CreateBound(New_xyz, New_xyz_A);

            double Move_X = New_x - oldPlace.X;
            double Move_Y = New_y - oldPlace.Y;
            double Rotate = New_Rotation - Old_Rotation;

            // Move the element to new location.
            XYZ new_xyz = new XYZ(Move_X, Move_Y, level.Elevation);

            //Start move using a transaction. 
            Transaction t = new Transaction(doc);
            t.Start("Move Element");

            ElementTransformUtils.MoveElement(doc, element.Id, new_xyz);
            ElementTransformUtils.RotateElement(doc, element.Id, New_Axis, Rotate);

            t.Commit();
        }

        //*****************************GetLevel()*****************************
        public Level GetLevel(Document doc, String levelName)
        {
            // LINQ to find the level by its name.
            Level level = (from lvl in new FilteredElementCollector(doc).
                               OfClass(typeof(Level)).
                               Cast<Level>()
                           where (lvl.Name == levelName)
                           select lvl).First();
            return level;
        }

        //*****************************GetParameterValue()*****************************
        /// get parameter's value
        public static string GetParameterValue(Parameter parameter)
        //public static Object GetParameterValue(Parameter parameter)
        {
            switch (parameter.StorageType)
            {
                case StorageType.Double:
                    //get value with unit, AsDouble() can get value without unit
                    return parameter.AsValueString();
                case StorageType.ElementId:
                    return parameter.AsElementId().IntegerValue.ToString();
                case StorageType.Integer:
                    //get value with unit, AsInteger() can get value without unit
                    return parameter.AsValueString();
                case StorageType.None:
                    return parameter.AsValueString();
                case StorageType.String:
                    return parameter.AsString();
                default:
                    return "";
            }
        }
    }
}