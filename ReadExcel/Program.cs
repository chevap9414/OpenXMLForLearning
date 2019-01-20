using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            ModelTypeImportExcel modelTypeImportExcel = new ModelTypeImportExcel();
            IModelTypeImportExcel ImodelTypeImportExcelService = modelTypeImportExcel;

            ImodelTypeImportExcelService.ImportExcel(@"Test.xlsx", "A");
            Console.Read();
        }

        private static List<List<string>> ReadExcelFile(string fileName)
        {
            fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Import\\" + fileName);
            string value = string.Empty;
            List<List<string>> rowValues = new List<List<string>>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                foreach (Sheet sheet in GetAllWorksheets(workbookPart))
                {
                    value = string.Empty;
                    WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    Worksheet worksheet = workbookPart.WorksheetParts.First().Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    rowValues.Add(new List<string> { "Sheet:" + sheet.Name });

                    List<Row> rows = sheetData.Descendants<Row>().ToList();
                    List<string> preRow = new List<string>();
                    List<string> cellValues = new List<string>();
                    
                    for(var i = 9; i < rows.Count; i++)
                    {
                        cellValues = new List<string>();
                        foreach (Cell cell in rows.ElementAt(i).Cast<Cell>())
                        {
                            string tempCellValue = GetCellValue(workbookPart, sheet, cell.CellReference);
                            if (cell.CellReference == "A" + (i+1))
                            {
                                if (string.IsNullOrEmpty(tempCellValue))
                                {
                                    tempCellValue = preRow[0];
                                }
                            }
                            if (cell.CellReference == "B" + (i + 1))
                            {
                                if (string.IsNullOrEmpty(tempCellValue))
                                {
                                    tempCellValue = preRow[1];
                                }
                            }
                            if (cell.CellReference == "C" + (i + 1))
                            {
                                if (string.IsNullOrEmpty(tempCellValue))
                                {
                                    tempCellValue = preRow[2];
                                }
                            }
                            if (cell.CellReference == "D" + (i + 1))
                            {
                                if (string.IsNullOrEmpty(tempCellValue))
                                {
                                    tempCellValue = preRow[3];
                                }
                            }
                            if (cell.CellReference == "E" + (i + 1))
                            {
                                if (string.IsNullOrEmpty(tempCellValue))
                                {
                                    tempCellValue = preRow[4];
                                }
                            }
                            cellValues.Add(tempCellValue);
                        }
                        preRow = cellValues;
                        rowValues.Add(cellValues);
                    }
                }
            }
            return rowValues;
        }

        private static string GetColumnName(string cellName)
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);
            return match.Value;
        }

        private static string GetCellValue(WorkbookPart workbookPart, Sheet sheet, string addressName)
        {
            string value = null;

            WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

            Cell theCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();
            if(theCell != null)
            {
                value = theCell.InnerText;
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            var stringSharedTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                            if(stringSharedTable != null)
                            {
                                value = stringSharedTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;
                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                case "1":
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }

            return value;
        }

        private static uint GetRowIndex(string cellName)
        {
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);
            return uint.Parse(match.Value);
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

        private static List<string> GetMergeCellValues(WorkbookPart workbookPart, Sheet sheet)
        {
            List<string> mergeCellvalues = new List<string>();
            WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
            {
                MergeCells mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().First();
                foreach (MergeCell mergeCell in mergeCells)
                {
                    string[] cellRef = mergeCell.Reference.Value.Split(':');
                    mergeCellvalues.Add(GetCellValue(workbookPart, sheet, cellRef[0]));
                }
            }
            return mergeCellvalues;
        }
    }
}
