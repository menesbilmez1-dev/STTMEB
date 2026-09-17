using System.Collections.Generic;

namespace sttmerve_cc.Models
{
    public class CategoryConfig
    {
        public string Name { get; set; } = string.Empty;
        public List<string> Types { get; set; } = new List<string>();
    }

    public class AppConfig
    {
        public List<CategoryConfig> Categories { get; set; } = new List<CategoryConfig>();
      
        public string ExcelFilePath { get; set; } = string.Empty;
    }
}

