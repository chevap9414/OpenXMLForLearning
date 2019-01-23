using DataLayer;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace ReadExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            ModelTypeImportExcel modelTypeImportExcel = new ModelTypeImportExcel();
            IModelTypeImportExcel ImodelTypeImportExcelService = modelTypeImportExcel;
            //ASHAOP_DEVEntities entities = new ASHAOP_DEVEntities();

            var uploadModel = new UploadFileImportModel
            {
                FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Import\\Test.xlsx")
            };
            if (!ValidateHeader(uploadModel))
            {
                
            }

            Console.Read();
        }



        public static int ImportExcel(UploadFileImportModel model)
        {
            if (!ValidateHeader(model))
            {
                return 7;
            }
            return ReadExcelOnThread(model);
        }

        public static bool ValidateHeader(UploadFileImportModel model)
        {
            bool IsSucceed = true;

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(model.SavePathSuccess, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                List<string> cellHeaderValueChecks = new List<string>() { "A6", "B6", "C6", "D6", "E6" };
                foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    foreach (string columnName in cellHeaderValueChecks)
                    {
                        if (string.IsNullOrEmpty(GetCellValue(workbookPart, sheet, columnName)))
                        {
                            return IsSucceed = false;
                        }
                    }
                }
            }
            return IsSucceed;
        }

        private static int ReadExcelOnThread(UploadFileImportModel model)
        {
            ModelTypeUploadModel value = new ModelTypeUploadModel();
            Thread thReadExcel = new Thread(() => { value = ReadExcel(model); });
            thReadExcel.Start();
            return value.UploadStatusID.HasValue ? 44 : 0;
        }


        private static ModelTypeUploadModel ReadExcel(UploadFileImportModel model)
        {
            ModelTypeUploadModel modelTypeUpload = new ModelTypeUploadModel();
            List<List<string>> rowValues = new List<List<string>>();
            string fileName = model.SavePathSuccess;
            List<ModelTypeTempSheetModel> sheetModels = new List<ModelTypeTempSheetModel>();
            ModelTypeTempSheetModel sheetModel;

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                int sheetCount = 0;
                foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                {
                    sheetModel = new ModelTypeTempSheetModel();
                    WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                    Worksheet worksheet = workbookPart.WorksheetParts.First().Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    List<Row> rows = sheetData.Descendants<Row>().ToList();
                    List<string> preRow = new List<string>();
                    List<string> cellValues = new List<string>();

                    int indexMainEquipStart = 0;
                    int indexPNoStart = 0;
                    int indexTypeStart = 0;
                    int indexVinStart = 0;
                    int indexEngineSerialNoStart = 0;
                    int indexErrorDescriptionStart = 0;
                    int inedexRowOfHeader_Start = 6;
                    int indexRowOfHeader_End = 8;

                    // Assign header value
                    foreach (Cell cell in rows.ElementAt(5).Descendants<Cell>())
                    {
                        if (new[] { "A6", "B6", "C6", "E6" }.Contains(cell.CellReference.Value))
                        {
                            string value = GetCellValue(workbookPart, sheet, cell.CellReference);
                            switch (cell.CellReference.Value)
                            {
                                case "A6":
                                    sheetModel.YM = value;
                                    break;
                                case "B6":
                                    sheetModel.Model = value;
                                    break;
                                case "C6":
                                    sheetModel.Door = value;
                                    break;
                                case "E6":
                                    sheetModel.Plant = value;
                                    break;
                            }
                        }
                    }
                    // Find column header index
                    for (var i = inedexRowOfHeader_Start; i < indexRowOfHeader_End; i++)
                    {
                        foreach (Cell cell in rows.ElementAt(i).Descendants<Cell>())
                        {
                            if (GetCellValue(workbookPart, sheet, cell.CellReference) == "MAIN EQUIPMENT")
                            {
                                indexMainEquipStart = GetColumnIndex(cell.CellReference);
                            }
                            if (GetCellValue(workbookPart, sheet, cell.CellReference) == "P.No.")
                            {
                                indexPNoStart = GetColumnIndex(cell.CellReference);
                            }
                            if (GetCellValue(workbookPart, sheet, cell.CellReference) == "TYPE")
                            {
                                indexTypeStart = GetColumnIndex(cell.CellReference);
                            }
                            if (GetCellValue(workbookPart, sheet, cell.CellReference) == "VIN")
                            {
                                indexVinStart = GetColumnIndex(cell.CellReference);
                            }
                            if (GetCellValue(workbookPart, sheet, cell.CellReference) == "ENGINE SERIAL No.")
                            {
                                indexEngineSerialNoStart = GetColumnIndex(cell.CellReference);
                            }
                            if (GetCellValue(workbookPart, sheet, cell.CellReference) == "Error Description")
                            {
                                indexErrorDescriptionStart = GetColumnIndex(cell.CellReference);
                            }
                        }
                    }

                    ModelTypeTempRowModel modelTypeTempRowModel;
                    ModelTypeTempEngineModel engineModel;
                    List<ModelTypeTempEquipmentModel> equipmentModels;
                    ModelTypeTempEquipmentModel equipmentModel;
                    List<ModelTypeTempTypeModel> typeModels;
                    ModelTypeTempTypeModel typeModel;

                    for (var i = 9; i < rows.Count; i++)
                    {
                        modelTypeTempRowModel = new ModelTypeTempRowModel();
                        equipmentModels = new List<ModelTypeTempEquipmentModel>();
                        typeModel = new ModelTypeTempTypeModel();
                        typeModels = new List<ModelTypeTempTypeModel>();
                        cellValues = new List<string>();
                        engineModel = new ModelTypeTempEngineModel();

                        modelTypeTempRowModel.RowNo = i + 1;
                        foreach (Cell cell in rows.ElementAt(i).Cast<Cell>())
                        {
                            string currentColumn = GetColumnName(cell.CellReference);
                            int currentIndex = GetColumnIndex(cell.CellReference);
                            string currentCellValue = GetCellValue(workbookPart, sheet, cell.CellReference);
                            int sequence = 1;

                            #region Engine
                            if (new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" }.Contains(GetColumnName(cell.CellReference)))
                            {

                                #region  Replace Value
                                if (cell.CellReference == "A" + (i + 1))
                                {
                                    if (string.IsNullOrEmpty(currentCellValue))
                                    {
                                        currentCellValue = preRow[0];
                                    }
                                }
                                if (cell.CellReference == "B" + (i + 1))
                                {
                                    if (string.IsNullOrEmpty(currentCellValue))
                                    {
                                        currentCellValue = preRow[1];
                                    }
                                }
                                if (cell.CellReference == "C" + (i + 1))
                                {
                                    if (string.IsNullOrEmpty(currentCellValue))
                                    {
                                        currentCellValue = preRow[2];
                                    }
                                }
                                if (cell.CellReference == "D" + (i + 1))
                                {
                                    if (string.IsNullOrEmpty(currentCellValue))
                                    {
                                        currentCellValue = preRow[3];
                                    }
                                }
                                if (cell.CellReference == "E" + (i + 1))
                                {
                                    if (string.IsNullOrEmpty(currentCellValue))
                                    {
                                        currentCellValue = preRow[4];
                                    }
                                }
                                #endregion

                                switch (GetColumnName(cell.CellReference))
                                {
                                    case "A":
                                        engineModel.SS = currentCellValue;
                                        break;
                                    case "B":
                                        engineModel.DISP = currentCellValue;
                                        break;
                                    case "C":
                                        engineModel.COMCARB = currentCellValue;
                                        break;
                                    case "D":
                                        engineModel.Grade = currentCellValue;
                                        break;
                                    case "E":
                                        engineModel.Mis = currentCellValue;
                                        break;
                                    case "F":
                                        engineModel.ModelCode01 = currentCellValue;
                                        break;
                                    case "G":
                                        engineModel.ModelCode02 = currentCellValue;
                                        break;
                                    case "H":
                                        engineModel.ModelCode03 = currentCellValue;
                                        break;
                                    case "I":
                                        engineModel.ModelCode04 = currentCellValue;
                                        break;
                                    case "J":
                                        engineModel.ModelCode05 = currentCellValue;
                                        break;

                                }

                            }
                            #endregion

                            #region MAIN EQUIPMENT
                            string columnEndGetEquipment = GetColumnName(GetMergeCellEndPosition(workbookPart, sheet, "K7"));
                            int indexMainEquipEnd = GetColumnIndex(columnEndGetEquipment);

                            if (currentIndex >= indexMainEquipStart && currentIndex <= indexMainEquipEnd) // Start K Column
                            {
                                equipmentModel = new ModelTypeTempEquipmentModel
                                {
                                    EquipmentName = GetCellValue(workbookPart, sheet, currentColumn + 9),
                                    EquipmentValue = currentCellValue,
                                    Sequence = sequence
                                };

                                sequence++;
                                equipmentModels.Add(equipmentModel);
                            }
                            #endregion

                            #region PNo
                            if (currentIndex == indexPNoStart)
                            {
                                modelTypeTempRowModel.PNo = currentCellValue;
                            }
                            #endregion

                            #region TYPE
                            if (currentIndex >= indexTypeStart && currentIndex <= indexVinStart - 1)
                            {
                                typeModel = new ModelTypeTempTypeModel
                                {
                                    ModelType = GetCellValue(workbookPart, sheet, currentColumn + 9),
                                    ModelCode = currentCellValue,
                                    Sequence = sequence
                                };
                                typeModels.Add(typeModel);
                            }
                            #endregion

                            #region VIN
                            if (currentIndex == indexVinStart)
                            {
                                modelTypeTempRowModel.VIN = currentCellValue;
                            }
                            #endregion

                            cellValues.Add(currentCellValue);
                        }
                        // End Cell
                        preRow = cellValues;
                        rowValues.Add(cellValues);
                        modelTypeTempRowModel.ModelTypeTempEngines.Add(engineModel);
                        modelTypeTempRowModel.ModelTypeTempEquipmentModels.AddRange(equipmentModels);
                        modelTypeTempRowModel.ModelTypeTempTypeModels.AddRange(typeModels);
                        sheetModel.ModelTypeTempRowModels.Add(modelTypeTempRowModel);
                    }
                    //End  Row
                    sheetModel.SheetNo = ++sheetCount;

                    modelTypeUpload.ModelTypeTempSheetModels.Add(sheetModel);
                }
            }
            return modelTypeUpload;
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

        private static uint GetRowIndex(string cellName)
        {
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);
            return uint.Parse(match.Value);
        }

        private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);
            if (row == null)
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
                foreach (MergeCell mergeCell in mergeCells.Descendants<MergeCell>())
                {
                    string[] cellRef = mergeCell.Reference.Value.Split(':');
                    mergeCellvalues.Add(GetCellValue(workbookPart, sheet, cellRef[0]));
                }
            }
            return mergeCellvalues;
        }

        private static string GetMergeCellEndPosition(WorkbookPart workbookPart, Sheet sheet, string addressStart)
        {
            string mergecellPosition = string.Empty;
            WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
            {
                MergeCells mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().First();
                foreach (MergeCell mergeCell in mergeCells.Descendants<MergeCell>())
                {
                    string[] cellMerge = mergeCell.Reference.Value.Split(':');
                    if (cellMerge[0] == addressStart)
                    {
                        mergecellPosition = cellMerge[1];
                    }
                }
            }
            return mergecellPosition;
        }

        private static int GetColumnIndex(string reference)
        {
            int ci = 0;
            reference = reference.ToUpper();
            for (int ix = 0; ix < reference.Length && reference[ix] >= 'A'; ix++)
                ci = (ci * 26) + ((int)reference[ix] - 64);
            return ci;
        }

        private static ModelTypeTempEngineModel GetEngineValue(string[] engineColumn)
        {
            ModelTypeTempEngineModel engine = new ModelTypeTempEngineModel();



            return engine;
        }
    }
}
