using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using sttmerve_cc.Models;

namespace sttmerve_cc.Services
{
    public class ExcelService
    {
        private string _filePath;
        private readonly CommandRepository _repository;
        public ExcelService(string filePath, CommandRepository repository)
        {
            _filePath = filePath;
            _repository = repository;
        }

        public void SetFilePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Excel file path is required.", nameof(filePath));
            _filePath = filePath;
        }

        public List<string> GetSheetNames()
        {
            try
            {
                using var workbook = new XLWorkbook(_filePath);
                return workbook.Worksheets.Select(w => w.Name).ToList();
            }
            catch (Exception ex)
            {
                throw new Exception($"Sayfa isimleri okunurken hata oluştu: {ex.Message}");
            }
        }

        public int GetColumnIndex(string sheetName, string mainHeader, string subHeader)
        {
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(sheetName);
            return FindColumnIndex(worksheet, mainHeader, subHeader);
        }

        private int FindColumnIndex(
            IXLWorksheet worksheet,
            string mainHeader,
            string subHeader)
        {
            int lastColumn = 0;
            var lastColRange = worksheet.LastColumnUsed();
            if (lastColRange != null) lastColumn = lastColRange.ColumnNumber();

            for (int col = 1; col <= lastColumn; col++)
            {
                string actualMainHeader = GetMergedCellValue(worksheet, 1, col);
                var row2Val = worksheet.Cell(2, col).Value.ToString()?.Trim() ?? string.Empty;

                if (string.Equals(
                        actualMainHeader.Trim(),
                        mainHeader.Trim(),
                        StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(
                        row2Val.Trim(),
                        subHeader.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }
            throw new Exception($"Şu sütun bulunamadı: {mainHeader} -> {subHeader}");
        }

        public int WriteCommandBatchToExcel(
            string sheetName,
            IReadOnlyList<CommandPayload> payloads,
            string? batchId = null)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                throw new ArgumentException("A worksheet must be configured.", nameof(sheetName));
            if (payloads.Count == 0)
                throw new ArgumentException("At least one command is required.", nameof(payloads));

            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(sheetName);
            var batchLedger = string.IsNullOrWhiteSpace(batchId)
                ? null
                : GetBatchLedger(workbook);
            if (batchLedger != null)
            {
                var existingRow = FindBatchRow(batchLedger, batchId!, sheetName);
                if (existingRow.HasValue) return existingRow.Value;
            }

            var commands = payloads
                .Select(payload => new
                {
                    Column = FindColumnIndex(
                        worksheet,
                        payload.Category.Trim(),
                        payload.Type.Trim()),
                    payload.Value
                })
                .ToList();

            if (commands.Select(command => command.Column).Distinct().Count() != commands.Count)
                throw new InvalidOperationException(
                    "A batch cannot write multiple values to the same Excel column.");

            var lastUsedColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            var lastUsedRow = worksheet.LastRowUsed()?.RowNumber() ?? 2;
            var targetRow = 3;
            for (; targetRow <= lastUsedRow; targetRow++)
            {
                var rowIsEmpty = true;
                for (var column = 1; column <= lastUsedColumn; column++)
                {
                    if (!string.IsNullOrWhiteSpace(
                            worksheet.Cell(targetRow, column).Value.ToString()))
                    {
                        rowIsEmpty = false;
                        break;
                    }
                }

                if (rowIsEmpty) break;
            }

            foreach (var command in commands)
            {
                var cell = worksheet.Cell(targetRow, command.Column);
                if (cell.HasFormula)
                    throw new InvalidOperationException(
                        $"Excel row {targetRow} contains a formula in a target column.");
                cell.Value = command.Value;
            }

            if (batchLedger != null)
            {
                var ledgerRow = (batchLedger.LastRowUsed()?.RowNumber() ?? 1) + 1;
                batchLedger.Cell(ledgerRow, 1).Value = batchId;
                batchLedger.Cell(ledgerRow, 2).Value = sheetName;
                batchLedger.Cell(ledgerRow, 3).Value = targetRow;
            }

            workbook.Save();
            return targetRow;
        }

        private IXLWorksheet GetBatchLedger(IXLWorkbook workbook)
        {
            const string ledgerName = "_STT_Batches";
            var ledger = workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals(sheet.Name, ledgerName, StringComparison.Ordinal));
            if (ledger == null)
            {
                ledger = workbook.Worksheets.Add(ledgerName);
                ledger.Cell(1, 1).Value = "BatchId";
                ledger.Cell(1, 2).Value = "SheetName";
                ledger.Cell(1, 3).Value = "ExcelRow";
            }
            else if (ledger.Cell(1, 1).GetString() != "BatchId" ||
                     ledger.Cell(1, 2).GetString() != "SheetName" ||
                     ledger.Cell(1, 3).GetString() != "ExcelRow")
            {
                throw new InvalidOperationException(
                    $"The reserved worksheet '{ledgerName}' has an unexpected layout.");
            }

            ledger.Visibility = XLWorksheetVisibility.VeryHidden;
            return ledger;
        }

        private int? FindBatchRow(
            IXLWorksheet ledger,
            string batchId,
            string sheetName)
        {
            var lastRow = ledger.LastRowUsed()?.RowNumber() ?? 1;
            for (var row = 2; row <= lastRow; row++)
            {
                if (!string.Equals(
                        ledger.Cell(row, 1).GetString(),
                        batchId,
                        StringComparison.Ordinal))
                {
                    continue;
                }

                if (!string.Equals(
                        ledger.Cell(row, 2).GetString(),
                        sheetName,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "A command batch ID cannot be reused for a different worksheet.");
                }

                if (!int.TryParse(ledger.Cell(row, 3).GetString(), out var excelRow))
                {
                    throw new InvalidDataException(
                        $"The stored Excel row for batch '{batchId}' is invalid.");
                }

                return excelRow;
            }

            return null;
        }

        private string GetMergedCellValue(IXLWorksheet worksheet, int row, int col)
        {
            var cell = worksheet.Cell(row, col);
            var range = worksheet.MergedRanges.FirstOrDefault(r => r.Contains(cell));
            if (range != null) return range.FirstCell().Value.ToString() ?? string.Empty;
            return cell.Value.ToString() ?? string.Empty;
        }

        public int GetFirstEmptyRow(string sheetName)
        {
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(sheetName);
            int lastRow = 0;
            var lastRowRange = worksheet.LastRowUsed();
            if (lastRowRange != null) lastRow = lastRowRange.RowNumber();

            int lastColumn = 0;
            var lastColRange = worksheet.LastColumnUsed();
            if (lastColRange != null) lastColumn = lastColRange.ColumnNumber();

            for (int row = 1; row <= lastRow + 1; row++)
            {
                bool isEmpty = true;
                for (int col = 1; col <= lastColumn; col++)
                {
                    if (!string.IsNullOrWhiteSpace(worksheet.Cell(row, col).Value.ToString()))
                    {
                        isEmpty = false;
                        break;
                    }
                }
                if (isEmpty) return row;
            }
            return lastRow + 1;
        }

        public void WriteValueAtRow(string sheetName, int row, int column, string value, bool saveImmediately = false)
        {
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(sheetName);
            worksheet.Cell(row, column).Value = value;
            if (saveImmediately) workbook.Save();
            else workbook.Save(); 
        }

        public DataTable ReadSheetData(string sheetName, string mainHeader, string subHeader)
        {
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(sheetName);
            var dt = new DataTable();
            dt.Columns.Add("ExcelRow", typeof(int));

            int lastColumn = 0;
            var lastColRange = worksheet.LastColumnUsed();
            if (lastColRange != null) lastColumn = lastColRange.ColumnNumber();

            int firstMatchCol = -1;
            int lastMatchCol = -1;
            var matchedColumns = new List<(int ColIndex, string ColumnName)>();

            // 1. Find the range of columns that match the criteria
            for (int col = 1; col <= lastColumn; col++)
            {
                string actualMainHeader = GetMergedCellValue(worksheet, 1, col).Trim();
                var row2Val = worksheet.Cell(2, col).Value.ToString()?.Trim() ?? string.Empty;

                if (string.Equals(actualMainHeader, mainHeader, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(row2Val, subHeader, StringComparison.OrdinalIgnoreCase))
                {
                    if (firstMatchCol == -1) firstMatchCol = col;
                    lastMatchCol = col;

                    // Construct unique name: {Main}_{Sub}_{Col}
                    string cleanMain = !string.IsNullOrEmpty(actualMainHeader) ? CleanColumnName(actualMainHeader) : "";
                    string cleanSub = !string.IsNullOrEmpty(row2Val) ? CleanColumnName(row2Val) : "";
                    string uniqueColName = BuildColumnName(cleanMain, cleanSub, col);

                    matchedColumns.Add((col, uniqueColName));
                }
            }

            if (firstMatchCol == -1) return dt;

            // 2. Add ALL columns in the range to the DataTable (including empty ones)
            for (int col = firstMatchCol; col <= lastMatchCol; col++)
            {
                string actualMainHeader = GetMergedCellValue(worksheet, 1, col).Trim();
                var row2Val = worksheet.Cell(2, col).Value.ToString()?.Trim() ?? string.Empty;

                string cleanMain = !string.IsNullOrEmpty(actualMainHeader) ? CleanColumnName(actualMainHeader) : "";
                string cleanSub = !string.IsNullOrEmpty(row2Val) ? CleanColumnName(row2Val) : "";

                string colName = BuildColumnName(cleanMain, cleanSub, col);

                var dataColumn = dt.Columns.Add(colName, typeof(string));
                dataColumn.ExtendedProperties["ExcelColumnIndex"] = col;
            }

            // 3. Find the maxRow (last row with data in any of the MATCHED columns)
            int maxRow = 0;
            int dataStartRow = 3;
            int lastRowUsed = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            foreach (var match in matchedColumns)
            {
                for (int row = dataStartRow; row <= lastRowUsed; row++)
                {
                    if (!string.IsNullOrWhiteSpace(worksheet.Cell(row, match.ColIndex).Value.ToString()))
                    {
                        if (row > maxRow) maxRow = row;
                    }
                }
            }

            // 4. Fill the DataTable
            if (maxRow >= dataStartRow)
            {
                for (int row = dataStartRow; row <= maxRow; row++)
                {
                    var dr = dt.NewRow();
                    dr["ExcelRow"] = row;
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (col.ColumnName == "ExcelRow") continue;
                        int excelCol = (int)col.ExtendedProperties["ExcelColumnIndex"]!;
                        dr[col.ColumnName] = worksheet.Cell(row, excelCol).Value.ToString() ?? string.Empty;
                    }
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }
        public void LoadSheet(string sheetName, out DataTable dataTable, out List<string> columns)
        {
            dataTable = new DataTable();
            columns = new List<string>();

            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(sheetName);

            // 1. Identify Columns from Row 1 (Main) and Row 2 (Sub)
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

            for (int col = 1; col <= lastColumn; col++)
            {
                string mainHeader = GetMergedCellValue(worksheet, 1, col);
                string subHeader = worksheet.Cell(2, col).Value.ToString()?.Trim() ?? string.Empty;

                string rawName = BuildColumnName(mainHeader, subHeader, col);
                string columnName = CleanColumnName(rawName);

                var dc = new DataColumn(columnName);

                dc.ExtendedProperties["ExcelColumnIndex"] = col;

                dataTable.Columns.Add(dc);
                columns.Add(columnName);
            }

            // Always add the tracking column for the Excel row number
            if (!dataTable.Columns.Contains("ExcelRow"))
            {
                dataTable.Columns.Add("ExcelRow", typeof(int));
            }

            // 2. Fill Data
            bool hasSubheaders = false;
            for (int col = 1; col <= lastColumn; col++)
            {
                if (!string.IsNullOrWhiteSpace(worksheet.Cell(2, col).Value.ToString()))
                {
                    hasSubheaders = true;
                    break;
                }
            }

            int startRow = hasSubheaders ? 3 : 2;
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            for (int row = startRow; row <= lastRow; row++)
            {
                var dr = dataTable.NewRow();
                dr["ExcelRow"] = row;

                for (int col = 1; col <= lastColumn; col++)
                {
                    var val = worksheet.Cell(row, col).Value.ToString() ?? string.Empty;

                    foreach (DataColumn c in dataTable.Columns)
                    {
                        if (c.ExtendedProperties.Contains("ExcelColumnIndex") &&
                            (int)c.ExtendedProperties["ExcelColumnIndex"] == col)
                        {
                            dr[c.ColumnName] = val;
                            break;
                        }
                    }
                }

                dataTable.Rows.Add(dr);
            }
        }

        public DataTable ReadConfigData(string sheetName, string subHeader)
        {
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(sheetName);
            var dt = new DataTable();
            dt.Columns.Add("ExcelRow", typeof(int));

            int lastColumn = 0;
            var lastColRange = worksheet.LastColumnUsed();
            if (lastColRange != null) lastColumn = lastColRange.ColumnNumber();

            int firstMatchCol = -1;
            int lastMatchCol = -1;
            var matchedColumns = new List<(int ColIndex, string ColumnName)>();

            // 1. Find the range of columns that match the subHeader
            for (int col = 1; col <= lastColumn; col++)
            {
                var row2Val = worksheet.Cell(2, col).Value.ToString()?.Trim() ?? string.Empty;
                if (string.Equals(row2Val, subHeader, StringComparison.OrdinalIgnoreCase))
                {
                    if (firstMatchCol == -1) firstMatchCol = col;
                    lastMatchCol = col;

                    string actualMainHeader = GetMergedCellValue(worksheet, 1, col).Trim();
                    string cleanMain = !string.IsNullOrEmpty(actualMainHeader) ? CleanColumnName(actualMainHeader) : "";
                    string cleanSub = !string.IsNullOrEmpty(row2Val) ? CleanColumnName(row2Val) : "";
                    string uniqueColName = BuildColumnName(cleanMain, cleanSub, col);

                    matchedColumns.Add((col, uniqueColName));
                }
            }

            if (firstMatchCol == -1) return dt;

            // 2. Add ALL columns in the range
            for (int col = firstMatchCol; col <= lastMatchCol; col++)
            {
                string actualMainHeader = GetMergedCellValue(worksheet, 1, col).Trim();
                var row2Val = worksheet.Cell(2, col).Value.ToString()?.Trim() ?? string.Empty;

                string cleanMain = !string.IsNullOrEmpty(actualMainHeader) ? CleanColumnName(actualMainHeader) : "";
                string cleanSub = !string.IsNullOrEmpty(row2Val) ? CleanColumnName(row2Val) : "";

                string colName = BuildColumnName(cleanMain, cleanSub, col);

                var dataColumn = dt.Columns.Add(colName, typeof(string));
                dataColumn.ExtendedProperties["ExcelColumnIndex"] = col;
            }

            // 3. Find maxRow
            int maxRow = 0;
            int dataStartRow = 3;
            int lastRowUsed = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            foreach (var match in matchedColumns)
            {
                for (int row = dataStartRow; row <= lastRowUsed; row++)
                {
                    if (!string.IsNullOrWhiteSpace(worksheet.Cell(row, match.ColIndex).Value.ToString()))
                    {
                        if (row > maxRow) maxRow = row;
                    }
                }
            }

            // 4. Fill
            if (maxRow >= dataStartRow)
            {
                for (int row = dataStartRow; row <= maxRow; row++)
                {
                    var dr = dt.NewRow();
                    dr["ExcelRow"] = row;
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (col.ColumnName == "ExcelRow") continue;
                        int excelCol = (int)col.ExtendedProperties["ExcelColumnIndex"]!;
                        dr[col.ColumnName] = worksheet.Cell(row, excelCol).Value.ToString() ?? string.Empty;
                    }
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        public DataTable GetSheetData(string sheetName, int startRow)
        {
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(sheetName);
            var dt = new DataTable();
            dt.Columns.Add("ExcelRow", typeof(int));

            int lastColumn = 0;
            var lastColRange = worksheet.LastColumnUsed();
            if (lastColRange != null) lastColumn = lastColRange.ColumnNumber();

            int firstCol = -1;
            int lastCol = -1;

            // 1. Find the range of columns that have ANY header info
            for (int col = 1; col <= lastColumn; col++)
            {
                string row1 = GetMergedCellValue(worksheet, 1, col).Trim();
                string row2 = worksheet.Cell(2, col).Value.ToString()?.Trim() ?? string.Empty;

                if (!string.IsNullOrEmpty(row1) || !string.IsNullOrEmpty(row2))
                {
                    if (firstCol == -1) firstCol = col;
                    lastCol = col;
                }
            }

            if (firstCol == -1) return dt;

            // 2. Add columns in range
            for (int col = firstCol; col <= lastCol; col++)
            {
                string row1 = GetMergedCellValue(worksheet, 1, col).Trim();
                string row2 = worksheet.Cell(2, col).Value.ToString()?.Trim() ?? string.Empty;

                string cleanMain = !string.IsNullOrEmpty(row1) ? CleanColumnName(row1) : "";
                string cleanSub = !string.IsNullOrEmpty(row2) ? CleanColumnName(row2) : "";

                string colName = BuildColumnName(cleanMain, cleanSub, col);

                var dataColumn = dt.Columns.Add(colName, typeof(string));
                dataColumn.ExtendedProperties["ExcelColumnIndex"] = col;
            }

            // 3. Find maxRow
            int maxRow = 0;
            int lastRowUsed = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            for (int row = startRow; row <= lastRowUsed; row++)
            {
                bool rowHasData = false;
                for (int col = 1; col <= lastColumn; col++)
                {
                    if (!string.IsNullOrWhiteSpace(worksheet.Cell(row, col).Value.ToString()))
                    {
                        rowHasData = true;
                        break;
                    }
                }
                if (rowHasData && row > maxRow) maxRow = row;
            }

            // 4. Fill
            if (maxRow >= startRow)
            {
                for (int row = startRow; row <= maxRow; row++)
                {
                    var dr = dt.NewRow();
                    dr["ExcelRow"] = row;
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (col.ColumnName == "ExcelRow") continue;
                        int excelCol = (int)col.ExtendedProperties["ExcelColumnIndex"]!;
                        dr[col.ColumnName] = worksheet.Cell(row, excelCol).Value.ToString() ?? string.Empty;
                    }
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        public void SaveSheetData(string sheetName, DataTable data)
        {
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(sheetName);

            foreach (DataRow row in data.Rows)
            {
                int excelRow = (int)row["ExcelRow"];
                foreach (DataColumn col in data.Columns)
                {
                    if (col.ColumnName == "ExcelRow") continue;
                    if (col.ExtendedProperties.Contains("ExcelColumnIndex"))
                    {
                        int excelCol = (int)col.ExtendedProperties["ExcelColumnIndex"]!;
                        var cell = worksheet.Cell(excelRow, excelCol);

                        if (cell.HasFormula) continue;

                        var val = row[col.ColumnName];
                        if (val != DBNull.Value && val != null)
                        {
                            cell.Value = ParseValue(val!.ToString());
                        }
                        else
                        {
                            cell.Value = string.Empty;
                        }
                    }
                }
            }
            workbook.Save();
        }

        public void WriteCommandToExcel(CommandPayload payload, string sheetName, string mainHeader)
        {
            try
            {
              
                int catCol = GetColumnIndex(sheetName, mainHeader, "Category");
                int typeCol = GetColumnIndex(sheetName, mainHeader, "Type");
                int valCol = GetColumnIndex(sheetName, mainHeader, "Value");

              
                int nextRow = GetFirstEmptyRow(sheetName);

                WriteValueAtRow(sheetName, nextRow, catCol, payload.Category, false);
                WriteValueAtRow(sheetName, nextRow, typeCol, payload.Type, false);
                WriteValueAtRow(sheetName, nextRow, valCol, payload.Value, true);

                Console.WriteLine($"[Excel] '{sheetName}' sayfasına {nextRow}. satır başarıyla yazıldı");
            }
            catch (Exception ex)
            {
              
                Console.WriteLine($"[Excel Hatası] Excel'e yazılamadı: {ex.Message}");
            }
        }
        private string BuildColumnName(string cleanMain, string cleanSub, int colIndex)
        {
            if (!string.IsNullOrEmpty(cleanMain) && !string.IsNullOrEmpty(cleanSub))
                return $"{cleanMain}_{cleanSub}_{colIndex}";

            if (!string.IsNullOrEmpty(cleanMain))
                return $"{cleanMain}_{colIndex}";

            if (!string.IsNullOrEmpty(cleanSub))
                return $"{cleanSub}_{colIndex}";

            return $"Boş_{colIndex}";
        }



        private string CleanColumnName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Bilinmeyen_Sütun";
            return name.Replace(" ", "_")
                       .Replace("/", "_")
                       .Replace("\\", "_")
                       .Replace(":", "_")
                       .Replace(".", "_")
                       .Replace("(", "_")
                       .Replace(")", "_");
        }

        private XLCellValue ParseValue(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            if (bool.TryParse(value, out bool boolValue))
                return boolValue;

            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleValueInvariant))
                return doubleValueInvariant;

            if (double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out double doubleValueCulture))
                return doubleValueCulture;

            return value;
        }
    }
}
