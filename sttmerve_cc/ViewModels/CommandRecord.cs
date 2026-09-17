using System;

namespace sttmerve_cc.Models
{
    public class CommandRecord
    {
        public long Id { get; set; } // Primary Key
        public string Category { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; }
        public bool IsExported { get; set; }
    }
}
