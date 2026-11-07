using System;

namespace sttmerve_cc.Models
{
    /// <summary>
    /// Represents the JSON payload sent by the Flutter app.
    /// </summary>
    public class CommandPayload
    {
        public string Category { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public string BatchId { get; set; } = string.Empty;
    }
}
