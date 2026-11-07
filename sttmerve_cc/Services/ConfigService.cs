using sttmerve_cc.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
using ClosedXML.Excel;

namespace sttmerve_cc.Services
{
    public class ConfigService
    {
        private AppConfig _config;
        private readonly string _configFilePath;

        public AppConfig Config => _config;
        public string ExcelFilePath => _config.ExcelFilePath;

        public ConfigService()
        {
            _configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
            Load();
        }
        public AppConfig GetConfig()
        {
            return _config;
        }

        public void Load()
        {
            if (File.Exists(_configFilePath))
            {
                try
                {
                    string json = File.ReadAllText(_configFilePath);
                    _config = JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
                }
                catch (Exception)
                {
                    _config = new AppConfig();
                }
            }
            else
            {
                _config = new AppConfig();
            }
        }

        public void Save()
        {
            try
            {
                string json = JsonSerializer.Serialize(_config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_configFilePath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Yapılandırma kaydedilirken hata oluştu: {ex.Message}");
            }
        }

        /// <summary>
        /// Synchronizes the application configuration schema (Categories and Types)
        /// directly from the Excel file's header structure.
        /// </summary>
        /// <param name="filePath">Path to the Excel file.</param>
        /// <param name="sheetName">Name of the worksheet.</param>
        public void SyncWithExcel(string filePath, string sheetName)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("Excel dosyası bulunamadı.");

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(sheetName);

            var newConfig = new AppConfig
            {
                ExcelFilePath = filePath,
                SheetName = sheetName
            };

            int lastColumn = 0;
            var lastColRange = worksheet.LastColumnUsed();
            if (lastColRange != null) lastColumn = lastColRange.ColumnNumber();

            for (int col = 1; col <= lastColumn; col++)
            {
                var cell1 = worksheet.Cell(1, col);
                string category = string.Empty;

                var range = worksheet.MergedRanges.FirstOrDefault(r => r.Contains(cell1));
                if (range != null)
                {
                    category = range.FirstCell().Value.ToString()?.Trim() ?? string.Empty;
                }
                else
                {
                    category = cell1.Value.ToString()?.Trim() ?? string.Empty;
                }

                string type = worksheet.Cell(2, col).Value.ToString()?.Trim() ?? string.Empty;

                if (!string.IsNullOrEmpty(type) && !string.IsNullOrEmpty(category))
                {
                    var existingCategory = newConfig.Categories.FirstOrDefault(c => c.Name == category);
                    if (existingCategory == null)
                    {
                        existingCategory = new CategoryConfig { Name = category, Types = new List<string>() };
                        newConfig.Categories.Add(existingCategory);
                    }

                    if (!existingCategory.Types.Contains(type))
                    {
                        existingCategory.Types.Add(type);
                    }
                }
                //It should be like this instead of && but i decided to keep both for now.
                else if (!string.IsNullOrEmpty(category) || !string.IsNullOrEmpty(type))
                {
                    var existingCategory = newConfig.Categories.FirstOrDefault(c => c.Name == category);
                    if (existingCategory == null)
                    {
                        existingCategory = new CategoryConfig { Name = category, Types = new List<string>() };
                        newConfig.Categories.Add(existingCategory);
                    }

                    if (!existingCategory.Types.Contains(type))
                    {
                        existingCategory.Types.Add(type);
                    }

                }
            }

            if (newConfig.Categories.Count > 0)
            {
                _config = newConfig;
            }
        }

        /// <summary>
        /// Updates the in-memory configuration based on the column headers of the active DataTable.
        /// Does NOT save to disk, allowing dynamic configuration per sheet.
        /// </summary>
        /// <param name="dt">The DataTable containing the Excel data.</param>
        public void SyncWithDataTable(DataTable dt)
        {
            if (dt == null || dt.Columns.Count == 0) return;

            var newConfig = new AppConfig();

            foreach (DataColumn col in dt.Columns)
            {
                string colName = col.ColumnName;
                if (colName == "ExcelRow") continue;

                // Expected format from ExcelService: Category_Type_Index (e.g., Temp_Sensor1_1)
                // We need to strip the index part (the last underscore part).
                int lastUnderscore = colName.LastIndexOf('_');
                if (lastUnderscore <= 0) continue;

                string nameWithoutIndex = colName.Substring(0, lastUnderscore);

                // Now nameWithoutIndex should be "Category_Type" (e.g., Temp_Sensor1)
                int splitIndex = nameWithoutIndex.LastIndexOf('_');

                string category = "";
                string type = "";

                if (splitIndex > 0)
                {
                    category = nameWithoutIndex.Substring(0, splitIndex);
                    type = nameWithoutIndex.Substring(splitIndex + 1);
                }
                else
                {
                    // It might just be "Category"
                    category = nameWithoutIndex;
                    type = "";
                }

                if (string.IsNullOrEmpty(category)) continue;

                var existingCategory = newConfig.Categories.FirstOrDefault(c => c.Name == category);
                if (existingCategory == null)
                {
                    existingCategory = new CategoryConfig { Name = category, Types = new List<string>() };
                    newConfig.Categories.Add(existingCategory);
                }

                if (!string.IsNullOrEmpty(type) && !existingCategory.Types.Contains(type))
                {
                    existingCategory.Types.Add(type);
                }
            }

            if (newConfig.Categories.Count > 0)
            {
                _config = newConfig;

            }
        }
    }
}
