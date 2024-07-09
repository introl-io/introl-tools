// See https://aka.ms/new-console-template for more information

using ClosedXML.Excel;
using Introl.Timesheets.Console.services;

Console.WriteLine("Hello, World!");

using var workbook = new XLWorkbook("./input/input_july.xlsx");

var sheetReader = new WorksheetReader();
var sheetWriter = new WorksheetWriter();
var inputModel = sheetReader.Process(workbook);
sheetWriter.Process(inputModel);
