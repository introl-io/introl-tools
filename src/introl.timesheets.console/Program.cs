// See https://aka.ms/new-console-template for more information

using ClosedXML.Excel;
using Introl.Timesheets.Console.services;

Console.WriteLine("Hello, World!");

using var workbook = new XLWorkbook("./input/input_july.xlsx");
var readerHelper = new WorksheetReaderHelper();
var writerHelper = new WorksheetWriterHelper();
var sheetReader = new WorksheetReader(readerHelper);
var sheetWriter = new WorksheetWriter(writerHelper);
var inputModel = sheetReader.Process(workbook);
sheetWriter.Process(inputModel);
