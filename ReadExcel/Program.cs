using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(ReadExcelFileSAX(@"Test1.xlsx"));
            Console.Read();
        }

        private static string ReadExcelFileSAX(string fileName)
        {
            fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Import\\" + fileName);
            string value = string.Empty;
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                foreach(Sheet sheet in GetAllWorksheets(workbookPart))
                {
                    WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    Worksheet worksheet = workbookPart.WorksheetParts.First().Worksheet;

                    Cell cell = GetCell(worksheet, "F", 9);
                    if(cell != null)
                    {
                        value = cell.InnerText;
                        if (cell.DataType != null)
                        {
                            switch (cell.DataType.Value)
                            {
                                case CellValues.SharedString:
                                    var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                    if(stringTable != null)
                                    {
                                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
            return value;
        }

        private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);
            if(row == null)
            {
                return null;
            }

            return row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0).First();
        }

        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        private static Sheets GetAllWorksheets(WorkbookPart workbookPart)
        {
            return workbookPart.Workbook.Sheets;
        }
    }
}
