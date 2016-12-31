using System;
using System.Data;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace OpenXmlAbstractions
{
    public class ExcelGenerationLibrary
    {
        // Populates a given start cell in a spreadsheet with a datatable
        public static void PopulateSpreadSheetWithDataTableTemplate(
            string docFilePath, 
            string tempFilePath, 
            string sheetName, 
            string startCell, 
            DataTable dt, 
            Dictionary<int, ColumnOptions> columnOptions)
        {
            if (!File.Exists(docFilePath))
            {
                throw new Exception("noFile");
            }

            // Deletes the temporary file if it exists
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
            // Copies the Template document to a specified temporary directory
            File.Copy(docFilePath, tempFilePath);

            try
            {
                using (var spreadSheet = SpreadsheetDocument.Open(tempFilePath, true))
                {
                    // Disable recalcuation before generating the document
                    //spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = false;

                    var startColLetter = StripNumbersFromString(startCell);
                    var startColNumber = ConvertColumnLetterToNumber(startColLetter);
                    var startRowNumber = Convert.ToUInt32(StripLettersFromString(startCell));
                    var col = startColNumber;

                    var worksheetPart = GetWorksheetPart(sheetName, spreadSheet);

                    for (int x = 0; x < dt.Rows.Count; x++)
                    {
                        for (int z = 0; z < dt.Columns.Count; z++)
                        {
                            ColumnOptions options = null;
                            columnOptions?.TryGetValue(col, out options);

                            InsertTextToCell(
                                spreadSheet, 
                                worksheetPart, 
                                dt.Rows[x][z].ToString(), 
                                col, 
                                startRowNumber,
                                options);

                            col++;
                        }
                        col = startColNumber;
                        startRowNumber++;
                    }

                    // Save the new worksheet.
                    worksheetPart.Worksheet.Save();

                    // Enable recalcuation again
                    //spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                }
            }
            catch (Exception)
            {
                File.Delete(tempFilePath);
                throw;
            }
        }

        private static WorksheetPart GetWorksheetPart(string sheetName, SpreadsheetDocument spreadSheet)
        {
            // Load the specified sheet
            var sheet = spreadSheet.WorkbookPart.Workbook
                .GetFirstChild<Sheets>()
                .Elements<Sheet>()
                .Where(s => s.Name == sheetName)
                .FirstOrDefault();

            if (sheet == null)
            {
                throw new Exception("noSheet");
            }

            return (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(sheet.Id.Value);
        }

        private static void InsertTextToCell(
            SpreadsheetDocument spreadSheet, 
            WorksheetPart worksheetPart, 
            string text, 
            int column, 
            uint row, 
            ColumnOptions options)
        {
            // If there is no text in the cell, return
            if (string.IsNullOrEmpty(text)) return;

            if (options == null)
            {
                options = new ColumnOptions();
            }

            // Converts the column number to the cooresponding excel column letter
            var actualColumn = ConvertColumnNumberToLetter(column);

            var cell = InsertCellInWorksheet(actualColumn, row, worksheetPart);

            var format = new CellFormat();
            format.Alignment = new Alignment();
            format.Alignment.WrapText = options.WrapText;
            format.Alignment.TextRotation = options.TextRotation;

            var font = new Font
            {
                Color = new Color { Rgb = new HexBinaryValue(options.TextColor) }
            };

            format.FontId = InsertFont(spreadSheet.WorkbookPart, font);

            if (options.TextColorChanges != null)
            {
                foreach (var changes in options.TextColorChanges)
                {
                    if (text.Contains(changes.Key))
                    {
                        var overrideFont = new Font
                        {
                            Color = new Color { Rgb = new HexBinaryValue(changes.Value) }
                        };

                        format.FontId = InsertFont(spreadSheet.WorkbookPart, overrideFont);
                    }
                }
            }

            cell.StyleIndex = InsertCellFormat(spreadSheet.WorkbookPart, format);

            if (options.TextReplacements != null)
            {
                foreach (var replacement in options.TextReplacements)
                {
                    text = text.Replace(replacement, string.Empty);
                }
            }

            int numResult;
            if (!int.TryParse(text, out numResult))
            {
                //// Get the SharedStringTablePart. If it does not exist, create a new one.
                //SharedStringTablePart shareStringPart;
                //if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                //{
                //    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                //}
                //else
                //{
                //    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                //}

                //// Insert the text into the SharedStringTablePart.
                //var index = InsertSharedStringItem(text, shareStringPart);

                //// Set the value of cell A1.
                //cell.CellValue = new CellValue(index.ToString());
                //cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                cell.CellValue = new CellValue(text);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
            else
            {
                cell.CellValue = new CellValue(text);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
        }

        private static string StripNumbersFromString(string input)
        {
            return Regex.Replace(input, @"\d", string.Empty);
        }

        private static string StripLettersFromString(string input)
        {
            return Regex.Replace(input, @"[a-zA-Z]", string.Empty);
        }

        private static string ConvertColumnNumberToLetter(int column)
        {
            column--;
            int remainder = column % 26; // figure out the right most leter
            string actualColumn = "" + (char)(remainder + 65); // ascii conversion

            column = column / 26;
            while (column > 0)
            {
                column--;
                remainder = column % 26;
                actualColumn = (char)(remainder + 65) + actualColumn;
                column = column / 26;
            }

            return actualColumn;
        }

        private static int ConvertColumnLetterToNumber(string columnLetter)
        {
            int columnNumber = 0;
            for (int i = columnLetter.Length - 1, exponent = 0; i >= 0; i--, exponent++)
            {
                columnNumber += ((int)columnLetter[i] - 64) * (int)Math.Pow(26, exponent);
            }

            return columnNumber;
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
                shareStringPart.SharedStringTable = new SharedStringTable();

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (var item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text) return i;

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            var row = sheetData.Elements<Row>()
                .Where(r => r.RowIndex == rowIndex)
                .FirstOrDefault();

            if (row == null)
            { 
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            var existingCell = row.Elements<Cell>()
                .Where(c => c.CellReference.Value == cellReference)
                .FirstOrDefault();

            // If there is not a cell with the specified column name, insert one.  
            if (existingCell != null)
            {
                return existingCell;
            }

            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;

            foreach (var cell in row.Elements<Cell>())
            {
                if (ColumnComparison(cell.CellReference.ToString(), cellReference))
                {
                    refCell = cell;
                    break;
                }
            }

            var newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            return newCell;
        }

        private static uint InsertCellFormat(WorkbookPart workbookPart, CellFormat cellFormat)
        {
            var cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().FirstOrDefault();
            if (cellFormats == null)
            {
                cellFormats = new CellFormats();
            }
            cellFormats.Append(cellFormat);
            return cellFormats.Count++;
        }

        private static uint InsertFont(WorkbookPart workbookPart, Font font)
        {
            var fonts = workbookPart.WorkbookStylesPart.Stylesheet.Elements<Fonts>().FirstOrDefault();
            if (fonts == null)
            {
                fonts = new Fonts();
            }
            fonts.Append(font);
            return fonts.Count++;
        }

        /// <summary>
        /// Takes two cell reference identfiers (such as "A1" and "B1") and returns true if 
        /// second param should go before first param.
        /// </summary>
        /// <param name="cellRefA"></param>
        /// <param name="cellRefB"></param>
        /// <returns></returns>
        private static bool ColumnComparison(string cellRefA, string cellRefB)
        {
            // first get the substring that is the column part
            int iA = 0;
            while (cellRefA[iA] >= 65) // while it isn't a number
                iA++;
            int iB = 0;
            while (cellRefB[iB] >= 65)
                iB++;

            string colA = cellRefA.Substring(0, iA);
            string colB = cellRefB.Substring(0, iB);

            if (colA.Length < colB.Length) // example C is shorter then AB
                return false; 

            // get the numeric column
            int numA = ConvertColumnLetterToNumber(colA);
            int numB = ConvertColumnLetterToNumber(colB);

            if (numA > numB) // example column D comes after column B
                return true;

            return false;
        }
    }
}