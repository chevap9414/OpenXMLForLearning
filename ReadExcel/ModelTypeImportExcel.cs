using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ReadExcel
{
    public class ModelTypeImportExcel : IModelTypeImportExcel
    {
        public ModelTypeImportExcel()
        {
        }

        public bool ImportExcel(string fileName, string uploadBy)
        {
            fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Import\\" + fileName);
            string value = string.Empty;
            List<List<string>> rowValues = new List<List<string>>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                {
                    value = string.Empty;
                    WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    Worksheet worksheet = workbookPart.WorksheetParts.First().Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    rowValues.Add(new List<string> { "Sheet:" + sheet.Name });

                    List<Row> rows = sheetData.Descendants<Row>().ToList();
                    List<string> preRow = new List<string>();
                    List<string> cellValues = new List<string>();

                    for (var i = 9; i < rows.Count; i++)
                    {
                        cellValues = new List<string>();
                        foreach (Cell cell in rows.ElementAt(i).Cast<Cell>())
                        {
                            string tempCellValue = GetCellValue(workbookPart, sheet, cell.CellReference);
                            if (cell.CellReference == "A" + (i + 1))
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
            return true;
        }

        private DataTable ReadExcel(string fileName)
        {
            DataTable data = new DataTable();

            return data;
        }

        private string GetCellValue(WorkbookPart workbookPart, Sheet sheet, string addressName)
        {
            string value = null;

            WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

            Cell theCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();
            if (theCell != null)
            {
                value = theCell.InnerText;
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            var stringSharedTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                            if (stringSharedTable != null)
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

        public bool IsPreverifyExcel(UploadFileImportModel model)
        {
            bool IsSucceed = true;

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(model.SavePathSuccess, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                List<string> cellHeaderValueChecks = new List<string>() { "A6", "B6", "C6", "D6", "E6" };
                foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    foreach(string columnName in cellHeaderValueChecks)
                    {
                        if(string.IsNullOrEmpty(GetCellValue(workbookPart, sheet, columnName)))
                        {
                            return IsSucceed = false;
                        }
                    }
                }
            }
            return IsSucceed;
        }
    }
}
