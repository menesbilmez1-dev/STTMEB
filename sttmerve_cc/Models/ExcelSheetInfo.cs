using System.Collections.Generic;

namespace sttmerve_cc.Models
{
    /// <summary>
    /// Represents the metadata for a worksheet within the Excel file.
    /// </summary>
    public class ExcelSheetInfo
    {
        public string Name { get; set; } = string.Empty;
        public int DataStartRow { get; set; }
    }
}
