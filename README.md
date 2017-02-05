# Project Name
Place isolator Revit family using an Excel file which is generated by AutoCAD.  The AutoCAD Excel file is generated using the command DATAEXTRACTION and getting all the XY coordinates and rotation of a AutoCAD block.  The Revit family I've build is a floor base hosted family.  You must pick the floor and level to generate the isolators. 

The plan for this project is to use it get points to locate base isolators in Revit through an Excel file.  The points generated my come from AutoCAD, Etabs or Rhino. 

Keywords: Revit, AutoCAD, Etabs, Rhino

## Installation
Copy .dll file and .addin to the [Revit Add-Ins folder](http://help.autodesk.com/view/RVT/2015/ENU/?guid=GUID-4FFDB03E-6936-417C-9772-8FC258A261F7).
Also look at my [DBLibrary](https://github.com/dannysbentley/DBLibrary) for missing methods. 

## Usage

Works in conjuntion with Revit. Developed with Revit API 2016.


## License

This sample is licensed under the terms of the [MIT License](https://opensource.org/licenses/MIT).

Copyright (c) 2017 Danny Bentley

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
