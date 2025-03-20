using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OutSystems.HubEdition.RuntimePlatform;

namespace OutSystems.NssCompareExcelFiles3
{


    public class CssCompareExcelFiles3 : IssCompareExcelFiles3
    {
        static byte[] FixInvalidColors(byte[] fileBytes)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(fileBytes, 0, fileBytes.Length); // Copy content to an expandable MemoryStream
                stream.Position = 0;

                using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
                {
                    WorkbookStylesPart stylesPart = document.WorkbookPart.WorkbookStylesPart;
                    if (stylesPart != null)
                    {
                        var stylesheet = stylesPart.Stylesheet;
                        bool modified = false;

                        foreach (var font in stylesheet.Fonts.Elements<Font>())
                        {
                            foreach (var color in font.Descendants<Color>())
                            {
                                if (color.Rgb != null && (color.Rgb.Value.Length != 8 || color.Rgb.Value == "0"))
                                {
                                    color.Rgb.Value = "FF000000"; // Default to black if invalid
                                    modified = true;
                                }
                            }
                        }

                        foreach (var fill in stylesheet.Fills.Elements<Fill>())
                        {
                            var fgColor = fill.PatternFill != null ? fill.PatternFill.ForegroundColor : null;
                            var bgColor = fill.PatternFill != null ? fill.PatternFill.BackgroundColor : null;
                            if (fgColor != null && (fgColor.Rgb == null || fgColor.Rgb.Value.Length != 8 || fgColor.Rgb.Value == "0"))
                            {
                                fgColor.Rgb = new HexBinaryValue("FFFFFFFF"); // Default to white if invalid
                                modified = true;
                            }
                            if (bgColor != null && (bgColor.Rgb == null || bgColor.Rgb.Value.Length != 8 || bgColor.Rgb.Value == "0"))
                            {
                                bgColor.Rgb = new HexBinaryValue("FFFFFFFF"); // Default to white if invalid
                                modified = true;
                            }
                        }

                        if (modified)
                        {
                            stylesPart.Stylesheet.Save();
                            Console.WriteLine("Fixed invalid colors in the file.");
                        }
                        else
                        {
                            Console.WriteLine("No invalid colors found in the file.");
                        }
                    }
                }

                return stream.ToArray(); // Return the modified byte array
            }
        }
        // Method to create a basic ClosedXML workbook and return its byte array
        private byte[] WarmUpClosedXML()
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("WarmUpSheet");
                    worksheet.Cell(1, 1).Value = "Warm-up";
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        // Main method to compare two Excel files
        public void MssCompare(byte[] ssExcelA, byte[] ssExcelB, string ssUniqueColumns, string ssSheetsToIgnoreString, string ssHeaderRowOfEachSheet, string ssTypeOfDocument, string ssWhichProcess, out byte[] ssResultExcel, out bool ssAreFilesDifferent, bool ssIsWarmUp)
        {
            ssResultExcel = new byte[] { };
            ssAreFilesDifferent = false;

            // Check for warm-up mode
            if (ssIsWarmUp)
            {
                ssResultExcel = WarmUpClosedXML();
                return;
            }
            try
            {
                ssExcelA = FixInvalidColors(ssExcelA);
                ssExcelB = FixInvalidColors(ssExcelB);
                // Parse input parameters
                string[] ssUniqueColsList = ssUniqueColumns.Split(',');
                int NULL_HASH_CODE = "".GetHashCode();
                const string BAMatrix = "BAMatrix";
                const string STANDARDISATION = "Standardisation";
                const string ACCESS_AREA_RESTRICTIONS = "Access Area Restrictions";
                int[] sheetsToIgnore = ssSheetsToIgnoreString == "" ? new int[0] : ssSheetsToIgnoreString.Split(',').Select(int.Parse).ToArray();
                int[] headerRows = ssHeaderRowOfEachSheet == "" ? new int[0] : ssHeaderRowOfEachSheet.Split(',').Select(int.Parse).ToArray();

                // Using memory streams to load the workbooks
                using (var stream1 = new MemoryStream(ssExcelA))
                using (var stream2 = new MemoryStream(ssExcelB))
                using (var streamResult = new MemoryStream())
                using (var workbook1 = new XLWorkbook(stream1))
                using (var workbook2 = new XLWorkbook(stream2))
                using (var workbookResult = new XLWorkbook())
                {
                    int numberOfSheets = workbook2.Worksheets.Count;
                    var sheetPairs = Enumerable.Range(0, numberOfSheets)
                                               .Select(sheetIndex => new { Sheet1 = workbook1.Worksheet(sheetIndex + 1), Sheet2 = workbook2.Worksheet(sheetIndex + 1), SheetIndex = sheetIndex })
                                               .ToList();

                    var differences = new List<bool>();

                    // Sequential execution for safety
                    foreach (var sheetPair in sheetPairs)
                    {
                        try
                        {
                            int headerRow = headerRows[sheetPair.SheetIndex];
                            var sheet1 = sheetPair.Sheet1;
                            var sheet2 = sheetPair.Sheet2;

                            // Ignore specified sheets
                            if (sheetsToIgnore.Contains(sheetPair.SheetIndex))
                            {
                                CopyWorksheet(sheet2, workbookResult);
                                continue;
                            }

                            var resultSheet = workbookResult.AddWorksheet(sheet2.Name);
                            //int maxColsSheet1 = GetLastUsedColumn(sheet1);
                            //int maxColsSheet2 = GetLastUsedColumn(sheet2);
                            //int maxColumns = Math.Max(maxColsSheet1, maxColsSheet2);

                            int maxColumns = GetMaxColumn(sheet1, sheet2);

                            CopyRowsToSheet(1, headerRow, maxColumns, sheet1, resultSheet);
                            bool sheetHasDifferences = false;
                            var headerRowMapExcel1 = GetHeaderRowMap(sheet1, headerRow);
                            var headerRowMapExcel2 = GetHeaderRowMap(sheet2, headerRow);
                            AddEmptyColumns(sheet1, sheet2, maxColumns, headerRow, headerRowMapExcel1, headerRowMapExcel2);

                            maxColumns = GetMaxColumn(sheet1, sheet2);

                            int resultRowCounter = headerRow;
                            // Compare the header row between the 2 excel files
                            ProcessHeaderRow(sheet1, sheet2, resultSheet, headerRow, maxColumns, ref resultRowCounter, headerRowMapExcel1, headerRowMapExcel2, ref sheetHasDifferences);

                            var rowMapExcel1 = GetRowMap(sheet1, headerRow, maxColumns, ssUniqueColsList);
                            var rowMapExcel2 = GetRowMap(sheet2, headerRow, maxColumns, ssUniqueColsList);
                            int rowToMoveToTheEnd = -1;

                            // Start after header row
                            for (int row = headerRow + 1; row <= Math.Max(GetLastUsedRow(sheet1), GetLastUsedRow(sheet2)); row++)
                            {
                                int row1Key = sheet1.Cell(row, 1).IsEmpty() ? "".GetHashCode() : GetRowKey(row, sheet1, maxColumns, ssUniqueColsList);
                                int rowNumInExcel2;

                                if (sheet1.Cell(row, 1).GetString().Contains(ACCESS_AREA_RESTRICTIONS))
                                {
                                    rowToMoveToTheEnd = row;
                                    continue;
                                }

                                if (rowMapExcel2.TryGetValue(row1Key, out rowNumInExcel2))
                                {
                                    ProcessRow(sheet1, sheet2, resultSheet, row, rowNumInExcel2, maxColumns, ref resultRowCounter, ref sheetHasDifferences);
                                }
                                else
                                {
                                    MarkRowAsDeleted(sheet1, resultSheet, row, maxColumns, ref resultRowCounter);
                                    sheetHasDifferences = true;
                                }

                                int row2Key = sheet2.Cell(row, 1).IsEmpty() ? "".GetHashCode() : GetRowKey(row, sheet2, maxColumns, ssUniqueColsList);
                                if (!rowMapExcel1.ContainsKey(row2Key))
                                {
                                    MarkRowAsAdded(sheet2, resultSheet, row, maxColumns, ref resultRowCounter);
                                    sheetHasDifferences = true;
                                }
                            }

                            AdjustColumnWidthsAndStyles(sheet1, resultSheet, headerRow, maxColumns);

                            // Move specific row to the end if conditions are met
                            if (rowToMoveToTheEnd != -1 && ssTypeOfDocument == BAMatrix && ssWhichProcess == STANDARDISATION && sheetPair.SheetIndex == 0)
                            {
                                MoveRowToTheEnd(sheet1, resultSheet, rowToMoveToTheEnd, resultRowCounter, maxColumns);
                            }

                            DeleteEmptyRows(resultSheet, headerRow, resultRowCounter);
                            
                            DeleteEmptyColumns(resultSheet, headerRow);

                            lock (differences)
                            {
                                differences.Add(sheetHasDifferences);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log("An error occurred processing sheet index " + sheetPair.SheetIndex + ": " + ex.Message);
                            Log("Stack Trace: " + ex.StackTrace);

                            throw; // Re-throw to stop further processing
                        }
                    }

                    workbookResult.SaveAs(streamResult);
                    ssResultExcel = streamResult.ToArray();
                    ssAreFilesDifferent = differences.Any(d => d);
                }
            }
            catch (Exception ex)
            {
                Log("An error occurred: " + ex.Message);
                Log("Stack Trace: " + ex.StackTrace);

                throw; // Re-throw the exception to be handled by the calling context
            }
        }

        // Adds empty columns in the sheets to ensure both have the same structure
        // Meaning that each column index will be compared on both excel files
        private void AddEmptyColumns(IXLWorksheet sheet1, IXLWorksheet sheet2, int maxColumns, int headerRow, Dictionary<int, int> headerRowMap1, Dictionary<int, int> headerRowMap2)
        {
            for (int column = 1; column <= maxColumns; column++)
            {
                string headerKey1String = sheet1.Cell(headerRow, column).GetString();
                string headerKey2String = sheet2.Cell(headerRow, column).GetString();

                int headerKey1 = GetCellKey(headerKey1String);
                int headerKey2 = GetCellKey(headerKey2String);

                if (!headerRowMap2.ContainsKey(headerKey1))
                {
                    sheet2.Column(column).InsertColumnsBefore(1);
                }

                if (!headerRowMap1.ContainsKey(headerKey2))
                {
                    // This is a rare case where it might happen that for example, excel1-column 4 doesn't exist in excel2 and column4-excel2 doesn't
                    // exist in excel1. If this happens, we have to increase the column by 2
                    if (!headerRowMap2.ContainsKey(headerKey1))
                    {
                        sheet1.Column(column).InsertColumnsAfter(1);
                        column++;
                    }
                    else
                    {
                        sheet1.Column(column).InsertColumnsBefore(1);
                    }
                }
            }
        }

        // Creates a dictionary mapping row keys to row numbers
        private Dictionary<int, int> GetRowMap(IXLWorksheet sheet, int headerRow, int maxColumns, string[] uniqueCols)
        {
            return Enumerable.Range(headerRow + 1, GetLastUsedRow(sheet) - headerRow)
                             .Select(row => new { Row = row, Key = GetRowKey(row, sheet, maxColumns, uniqueCols) })
                             .Where(x => x.Key != "".GetHashCode())
                             .ToDictionary(x => x.Key, x => x.Row);
        }

        // Creates a dictionary mapping column keys to column numbers
        private Dictionary<int, int> GetHeaderRowMap(IXLWorksheet sheet, int headerRow)
        {
            return Enumerable.Range(1, GetLastUsedColumn(sheet))
                             .Select(col => new { Col = col, Key = GetCellKey(sheet.Cell(headerRow, col).GetString()) })
                             .Where(x => x.Key != "".GetHashCode())
                             .ToDictionary(x => x.Key, x => x.Col);
        }

        // Generates a key for a row based on the unique columns
        private int GetRowKey(int rowNum, IXLWorksheet sheet, int maxColumns, string[] uniqueCols)
        {
            return string.Join("", uniqueCols.Select(col => sheet.Cell(rowNum, int.Parse(col)).GetString())).GetHashCode();
        }

        // Generates a key for a cell value
        private int GetCellKey(string cellValue)
        {
            return cellValue.GetHashCode();
        }

        // Processes a single row, comparing cells and marking differences
        private void ProcessRow(IXLWorksheet sheet1, IXLWorksheet sheet2, IXLWorksheet resultSheet, int row1, int row2, int maxColumns, ref int resultRowCounter, ref bool sheetHasDifferences)
        {
            for (int column = 1; column <= maxColumns; column++)
            {
                var cell1 = sheet1.Cell(row1, column);
                var cell2 = sheet2.Cell(row2, column);
                var resultCell = resultSheet.Cell(resultRowCounter, column);

                if (cell1.IsEmpty() && cell2.IsEmpty()) continue;

                if (cell1.IsEmpty())
                {
                    resultCell.Value = cell2.Value;
                    resultCell.Style.Font.FontColor = XLColor.Green;
                    resultCell.Style.Font.Underline = XLFontUnderlineValues.Single;
                    sheetHasDifferences = true;
                }
                else if (cell2.IsEmpty())
                {
                    resultCell.Value = cell1.Value;
                    resultCell.Style.Font.FontColor = XLColor.Red;
                    resultCell.Style.Font.Strikethrough = true;
                    sheetHasDifferences = true;
                }
                else if (!cell1.GetString().Equals(cell2.GetString()))
                {
                    var richText = resultCell.GetRichText();
                    richText.ClearText();
                    var part1 = richText.AddText(cell1.GetString());
                    part1.SetFontColor(XLColor.Red);
                    part1.SetStrikethrough();
                    var part2 = richText.AddText(cell2.GetString());
                    part2.SetFontColor(XLColor.Green);
                    part2.SetUnderline();
                    sheetHasDifferences = true;
                }
                else
                {
                    resultCell.Value = cell1.Value;
                }
            }
            resultRowCounter++;
        }

        // Processes the header row, comparing cells and marking differences
        private void ProcessHeaderRow(IXLWorksheet sheet1, IXLWorksheet sheet2, IXLWorksheet resultSheet, int headerRow, int maxColumns, ref int resultRowCounter, Dictionary<int, int> headerRowMap1, Dictionary<int, int> headerRowMap2, ref bool sheetHasDifferences)
        {
            int rowNumInExcel1 = headerRow;
            int rowNumInExcel2 = headerRow;
            string cellValueExcel1 = sheet1.Cell(headerRow, 1).GetString();
            int row1Key = GetCellKey(cellValueExcel1);

            if (headerRowMap2.ContainsKey(row1Key))
            {
                ProcessRow(sheet1, sheet2, resultSheet, rowNumInExcel1, rowNumInExcel2, maxColumns, ref resultRowCounter, ref sheetHasDifferences);
            }
            else
            {
                MarkRowAsDeleted(sheet1, resultSheet, headerRow, maxColumns, ref resultRowCounter);
                sheetHasDifferences = true;
            }

            string cellValueExcel2 = sheet2.Cell(headerRow, 1).GetString();
            int row2Key = GetCellKey(cellValueExcel2);
            if (!headerRowMap1.ContainsKey(row2Key))
            {
                MarkRowAsAdded(sheet2, resultSheet, headerRow, maxColumns, ref resultRowCounter);
                sheetHasDifferences = true;
            }
        }

        // Marks a row as deleted by applying a specific style
        private void MarkRowAsDeleted(IXLWorksheet sourceSheet, IXLWorksheet resultSheet, int row, int maxColumns, ref int resultRowCounter)
        {
            for (int column = 1; column <= maxColumns; column++)
            {
                var resultCell = resultSheet.Cell(resultRowCounter, column);
                resultCell.Style.Font.FontColor = XLColor.Red;
                resultCell.Style.Font.Strikethrough = true;
                resultCell.Value = sourceSheet.Cell(row, column).Value;
            }
            resultRowCounter++;
        }

        // Marks a row as added by applying a specific style
        private void MarkRowAsAdded(IXLWorksheet sourceSheet, IXLWorksheet resultSheet, int row, int maxColumns, ref int resultRowCounter)
        {
            for (int column = 1; column <= maxColumns; column++)
            {
                var resultCell = resultSheet.Cell(resultRowCounter, column);
                resultCell.Style.Font.FontColor = XLColor.Green;
                resultCell.Style.Font.Underline = XLFontUnderlineValues.Single;
                resultCell.Value = sourceSheet.Cell(row, column).Value;
            }

            // Handle new columns in sourceSheet
            for (int column = maxColumns + 1; column <= GetLastUsedColumn(sourceSheet); column++)
            {
                var cell = sourceSheet.Cell(row, column);
                var resultCell = resultSheet.Cell(resultRowCounter, column);
                resultCell.Value = cell.Value;
                resultCell.Style.Font.FontColor = XLColor.Green;
                resultCell.Style.Font.Underline = XLFontUnderlineValues.Single;
            }

            resultRowCounter++;
        }

        // Adjusts column widths and styles in the result sheet to match the source sheet
        private void AdjustColumnWidthsAndStyles(IXLWorksheet sourceSheet, IXLWorksheet resultSheet, int headerRow, int maxColumns)
        {
            for (int column = 1; column <= maxColumns; column++)
            {
                var resultCell = resultSheet.Cell(headerRow, column);
                resultSheet.Column(column).Width = sourceSheet.Column(column).Width;
                resultCell.Style.Font.Bold = sourceSheet.Cell(headerRow, column).Style.Font.Bold;
            }

            // Handle new columns
            for (int column = maxColumns + 1; column <= GetLastUsedColumn(sourceSheet); column++)
            {
                resultSheet.Column(column).Width = sourceSheet.Column(column).Width;
                var resultCell = resultSheet.Cell(headerRow, column);
                resultCell.Style.Font.Bold = true;
                resultCell.Style.Font.FontColor = XLColor.Green;
            }
        }

        // Copies the contents of a worksheet to another workbook
        private void CopyWorksheet(IXLWorksheet sourceSheet, XLWorkbook destinationWorkbook)
        {
            var destinationSheet = destinationWorkbook.AddWorksheet(sourceSheet.Name);
            foreach (var cell in sourceSheet.CellsUsed())
            {
                var destinationCell = destinationSheet.Cell(cell.Address);
                destinationCell.Value = cell.Value;
                destinationCell.Style = cell.Style;
            }
        }

        // Gets the number of the last used row in a sheet
        private int GetLastUsedRow(IXLWorksheet sheet)
        {
            return sheet.LastRowUsed().RowNumber();
        }

        // Gets the number of the last used column in a sheet
        private int GetLastUsedColumn(IXLWorksheet sheet)
        {
            return sheet.LastColumnUsed().ColumnNumber();
        }

        private int GetMaxColumn(IXLWorksheet sheet1, IXLWorksheet sheet2)
        {
            int maxColsSheet1 = GetLastUsedColumn(sheet1);
            int maxColsSheet2 = GetLastUsedColumn(sheet2);
            return Math.Max(maxColsSheet1, maxColsSheet2);

        }

        // Copies a range of rows from one sheet to another
        private void CopyRowsToSheet(int fromRow, int untilRow, int maxColumns, IXLWorksheet srcSheet, IXLWorksheet dstSheet)
        {
            for (int row = fromRow; row < untilRow; row++)
            {
                for (int column = 1; column <= maxColumns; column++)
                {
                    var srcCell = srcSheet.Cell(row, column);
                    var dstCell = dstSheet.Cell(row, column);
                    dstCell.Value = srcCell.Value;
                    dstCell.Style = srcCell.Style;
                }
            }
        }

        // Deletes rows in a sheet that are empty
        private void DeleteEmptyRows(IXLWorksheet sheet, int startRow, int endRow)
        {
            for (int row = endRow; row >= startRow; row--)
            {
                if (sheet.Row(row).CellsUsed().All(cell => cell.IsEmpty()))
                {
                    sheet.Row(row).Delete();
                }
            }
        }

        private void DeleteEmptyColumns(IXLWorksheet sheet, int headerRow)
        {
            //idetify the first and last columns in use
            var firstCol = sheet.FirstColumnUsed().ColumnNumber();
            var lastCol = sheet.LastColumnUsed().ColumnNumber();

            //iterate from right to left so column indices remain valid after deletions
            for (int col = lastCol; col >= firstCol; col--) 
            {
                var headerCell = sheet.Cell(headerRow, col);
                if (headerCell.IsEmpty())
                {
                    sheet.Column(col).Delete();
                }
            }
        }

            // Moves a specific row to the end of the result sheet
            private void MoveRowToTheEnd(IXLWorksheet srcSheet, IXLWorksheet resultSheet, int srcRow, int lastUsedRow, int maxColumns)
        {
            for (int column = 1; column <= maxColumns; column++)
            {
                var srcCell = srcSheet.Cell(srcRow, column);
                var dstCell = resultSheet.Cell(lastUsedRow, column);
                dstCell.Value = srcCell.Value;
                dstCell.Style = srcCell.Style;
            }
            resultSheet.Row(srcRow).Delete();
        }

        // Logs the messages into Outsystems
        private void Log(string message)
        {
            const int MaxLogLength = 1000;
            int messageLength = message.Length;

            for (int i = 0; i < messageLength; i += MaxLogLength)
            {
                string logPart = message.Substring(i, Math.Min(MaxLogLength, messageLength - i));
                GenericExtendedActions.LogMessage(AppInfo.GetAppInfo().OsContext, logPart, "CompareExcelFiles3");
            }
        }

        //private void Log(string message)
        //{
        //    Console.WriteLine("[INFO]: " + message);
        //}

        // Utility function used for dev purposes. You can call it like this:
        private void SaveExcelFiles(XLWorkbook workbook1, XLWorkbook workbook2, string filePath1, string filePath2)
        {
            using (var stream1 = new MemoryStream())
            using (var stream2 = new MemoryStream())
            {
                workbook1.SaveAs(stream1);
                workbook2.SaveAs(stream2);

                File.WriteAllBytes(filePath1, stream1.ToArray());
                File.WriteAllBytes(filePath2, stream2.ToArray());
            }
        }
    }
}