using System;
using System.IO;
using System.Text.Json;

namespace sttmerve_cc.Services
{
    public class AppSettings
    {
        public string ExcelFilePath { get; set; } = string.Empty;
    }

    public class SettingsService
    {
        private readonly string _settingsFilePath;

        public SettingsService()
        {
            _settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.json");
        }

        public AppSettings LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    string json = File.ReadAllText(_settingsFilePath);
                    return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ayarlar yüklenirken hata oluştu: {ex.Message}");
            }
            return new AppSettings();
        }

        public void SaveSettings(AppSettings settings)
        {
            try
            {
                string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_settingsFilePath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ayarlar kaydedilirken hata oluştu: {ex.Message}");
            }
        }
    }
}
