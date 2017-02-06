using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.Diagnostics;

namespace SOM.RevitTools.PlaceIsolator
{
    class ExcelReader
    {
        //*****************************Get_ExcelFileToList()*****************************
        public List<IsoObj> Get_ExcelFileToList(string Etabsfile_path)
        {
            List<IsoObj> List_Isolators = new List<IsoObj>();
            try
            {

                // Gets Excel file using a file path 
                var package = new ExcelPackage(new FileInfo(Etabsfile_path));
                // Count the number of Worksheets are in the file.
                int num = package.Workbook.Worksheets.Count;
                // get first workshet 
                ExcelWorksheet WorkSheets = package.Workbook.Worksheets[1];
                //check name
                string name = WorkSheets.Name;
                var End_WorkSheets = WorkSheets.Dimension.End;
                // Starts on the second row avoiding header. 
                for (int row = 2; row <= End_WorkSheets.Row; row++)
                {
                    IsoObj objIsolator = new IsoObj();
                    for (int col = 1; col <= End_WorkSheets.Column; col++)
                    {
                        if (col == 2)
                            objIsolator.X = WorkSheets.Cells[row, col].Text;
                        if (col == 3)
                            objIsolator.Y = WorkSheets.Cells[row, col].Text;
                        //if (col == 4)
                            //objIsolator.R = WorkSheets.Cells[row, col].Text;
                        if (col == 5)
                            objIsolator.ID = WorkSheets.Cells[row, col].Text;
                    }
                    List_Isolators.Add(objIsolator);
                }
            }
            catch (Exception e)
            {
                string x = e.Message;
            }
            return List_Isolators;
        }
    }
}
