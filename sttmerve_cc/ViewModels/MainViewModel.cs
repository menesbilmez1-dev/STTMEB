using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using sttmerve_cc.Helpers;
using sttmerve_cc.Models;
using sttmerve_cc.Services;
using System.Data;

namespace sttmerve_cc.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private ExcelService? _excelService;
        private readonly HttpServerService _httpServerService;
        private readonly SettingsService _settingsService;
        private readonly ConfigService _configService;
        private readonly CommandRepository _repository;


        private string _excelFilePath = string.Empty;
        private bool _isServerRunning;
        private ExcelSheetInfo? _selectedSheet;
        private DataTable? _sheetDataTable;

        public ICommand SaveToExcelCommand { get; }
        public ICommand LoadExcelCommand { get; }
        public ICommand LoadExcelFromCurrentPathCommand { get; }
        public ICommand SaveConfigFromExcelCommand { get; }
        public ICommand SyncWithExcelCommand { get; }
        public ICommand StartServerCommand { get; }
        public ICommand StopServerCommand { get; }

        private string _statusMessage = "Hazır";

        public ObservableCollection<ExcelSheetInfo> Sheets { get; } = new();
        public ObservableCollection<string> Logs { get; } = new();

        public string ExcelFilePath
        {
            get => _excelFilePath;
            set => SetProperty(ref _excelFilePath, value);
        }

        public bool IsServerRunning
        {
            get => _isServerRunning;
            set => SetProperty(ref _isServerRunning, value);
        }

        public ExcelSheetInfo? SelectedSheet
        {
            get => _selectedSheet;
            set
            {
                if (SetProperty(ref _selectedSheet, value))
                {
                    LoadSheet();
                }
            }
        }
        private string _serverUrl = string.Empty;


        public string ServerUrl
        {
            get => _serverUrl;
            set => SetProperty(ref _serverUrl, value);
        }


        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public DataTable? SheetDataTable
        {
            get => _sheetDataTable;
            set => SetProperty(ref _sheetDataTable, value);
        }

        public MainViewModel(
            ExcelService excelService,
            HttpServerService httpServerService,
            CommandRepository commandRepository,
            ConfigService configService)
        {
            _repository = commandRepository;
            _settingsService = new SettingsService();
            _configService = configService;
            _excelService = excelService;
            _httpServerService = httpServerService;

            // Bind events
            _httpServerService.CommandReceived += OnCommandReceived;
            _httpServerService.LogMessageReceived += OnLogReceived;
            _httpServerService.ConnectionStatusChanged += OnConnectionStatusChanged;


            _httpServerService.ServerStarted += (url) =>
            {

                System.Windows.Application.Current.Dispatcher.Invoke(() => ServerUrl = url);
            };

            SaveToExcelCommand = new RelayCommand(_ => SaveToExcel());
            LoadExcelCommand = new RelayCommand(_ => LoadExcel());
            SyncWithExcelCommand = new RelayCommand(_ => SyncWithExcel());

            SaveConfigFromExcelCommand = new RelayCommand(_ => SaveConfigFromExcel());

            LoadExcelFromCurrentPathCommand = new RelayCommand(_ => LoadExcelFromCurrentPath());
            StartServerCommand = new RelayCommand(_ => StartServer(), _ => !IsServerRunning);
            StopServerCommand = new RelayCommand(_ => StopServer(), _ => IsServerRunning);

            // Load last used Excel file
            var settings = _settingsService.LoadSettings();
            if (!string.IsNullOrWhiteSpace(settings.ExcelFilePath) && System.IO.File.Exists(settings.ExcelFilePath))
            {
                LoadExcelFromFile(settings.ExcelFilePath);
            }
        }
        private void AddLog(string message)
        {
            string logEntry = $"[{DateTime.Now:HH:mm:ss}] {message}";
            Logs.Insert(0, logEntry);
            if (Logs.Count > 100) Logs.RemoveAt(100);
        }

        private void LoadExcel()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Dosyaları (*.xlsx)|*.xlsx|Tüm Dosyalar (*.*)|*.*",
                Title = "Excel Dosyası Seç"
            };

            if (dialog.ShowDialog() == true)
            {
                LoadExcelFromFile(dialog.FileName);
            }
        }

        private void LoadExcelFromCurrentPath()
        {
            if (!string.IsNullOrWhiteSpace(ExcelFilePath) && System.IO.File.Exists(ExcelFilePath))
            {
                LoadExcelFromFile(ExcelFilePath);
            }
            else
            {
                AddLog("Geçersiz dosya yolu. Lütfen önce 'Gözat...' ile bir dosya seçin veya yolu doğru yazın.");
                StatusMessage = "Geçersiz dosya yolu";
            }
        }

        private void LoadExcelFromFile(string path)
        {
            try
            {
                ExcelFilePath = path;
                _excelService!.SetFilePath(ExcelFilePath);
                // Removed _configService.SyncWithExcel(path) to eliminate Config sheet dependency

                AddLog($"Excel Yükleniyor: {ExcelFilePath}");
                var sheetNames = _excelService.GetSheetNames();
                Sheets.Clear();
                foreach (var name in sheetNames)
                {
                    Sheets.Add(new ExcelSheetInfo { Name = name, DataStartRow = 3 });
                }
                StatusMessage = "Excel başarıyla yüklendi";
                AddLog("Excel başarıyla yüklendi.");

                // Save the settings so it loads next time
                _settingsService.SaveSettings(new AppSettings { ExcelFilePath = path });
            }
            catch (Exception ex)
            {
                AddLog($"Excel yüklenirken hata oluştu: {ex.Message}");
                StatusMessage = "Excel yükleme hatası";
            }
        }

        // Inside MainViewModel.cs

        private void LoadSheet()
        {
            if (SelectedSheet == null) return;

            try
            {
                AddLog($"Sayfa yükleniyor: {SelectedSheet.Name}");

                // 1. Load the structure from Excel (Existing logic)
                _excelService.LoadSheet(SelectedSheet.Name, out var dataTable, out var columns);
                this.SheetDataTable = dataTable;

                // 2. THE SAFETY NET RECOVERY: 
                // Pull all data from SQLite and inject it into the DataTable
                var history = _repository.GetAll();

                foreach (var record in history)
                {
                    // We only inject records that belong to this specific sheet/category
                    // (Assuming your sheet name or category matches)
                    if (record.Category == SelectedSheet.Name || record.Category == "All")
                    {
                        InjectRecordToDataTable(record);
                    }
                }

                AddLog("Veriler Veritabanından geri yüklendi.");
            }
            catch (Exception ex)
            {
                AddLog($"Hata: {ex.Message}");
            }
        }

        private void InjectRecordToDataTable(CommandPayload payload)
        {
            if (SheetDataTable == null) return;

            // Find the correct column index using your existing ExcelService logic
            int excelCol = _excelService.GetColumnIndex(SelectedSheet!.Name, payload.Category, payload.Type);

            // Find the matching DataColumn in the DataTable
            DataColumn? targetCol = null;
            foreach (DataColumn col in SheetDataTable.Columns)
            {
                if (col.ExtendedProperties.Contains("ExcelColumnIndex") &&
                    (int)col.ExtendedProperties["ExcelColumnIndex"] == excelCol)
                {
                    targetCol = col;
                    break;
                }
            }

            if (targetCol != null)
            {
                // Find next row or create new
                int nextRow;
                if (SheetDataTable.Rows.Count > 0)
                    nextRow = SheetDataTable.AsEnumerable().Max(r => (int)r["ExcelRow"]) + 1;
                else
                    nextRow = SelectedSheet!.DataStartRow;

                DataRow newRow = SheetDataTable.NewRow();
                newRow["ExcelRow"] = nextRow;
                newRow[targetCol.ColumnName] = payload.Value;
                SheetDataTable.Rows.Add(newRow);
            }
        }


        private void SaveToExcel()
        {
            if (_excelService == null || SelectedSheet == null || SheetDataTable == null)
            {
                AddLog("Hata: Excel dosyası yüklenmedi veya geçerli veri yok.");
                return;
            }

            try
            {
                AddLog($"Değişiklikler {SelectedSheet.Name} dosyasına kaydediliyor...");
                _excelService.SaveSheetData(SelectedSheet.Name, SheetDataTable);
                AddLog("Kaydetme başarılı.");
                StatusMessage = "Kaydetme Başarılı";
            }
            catch (Exception ex)
            {
                AddLog($"Kaydetme hatası: {ex.Message}");
                StatusMessage = "Kaydetme Başarısız";
            }
        }

        private void StartServer()
        {
            try
            {
                if (_excelService == null || SelectedSheet == null)
                {
                    AddLog("Hata: Sunucuyu başlatmadan önce Excel dosyası ve sayfası seçilmelidir.");
                    StatusMessage = "Excel sayfası seçilmedi";
                    return;
                }

                _configService.SyncWithExcel(ExcelFilePath, SelectedSheet.Name);
                _configService.Save();
                _excelService.SetFilePath(ExcelFilePath);
                _httpServerService.Start();
                IsServerRunning = true;
                StatusMessage = "Sunucu Çalışıyor (Port 8080)";
                AddLog("Sunucu başlatıldı.");
            }
            catch (Exception ex)
            {
                AddLog($"Sunucu başlatılırken hata oluştu: {ex.Message}");
                StatusMessage = "Sunucu başlatılamadı";
            }
        }

        private void StopServer()
        {
            _httpServerService.Stop();
            IsServerRunning = false;
            StatusMessage = "Sunucu Durduruldu";
            AddLog("Sunucu durduruldu.");
        }


        private void OnCommandReceived(List<CommandPayload> payloads)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                if (_excelService == null || SelectedSheet == null)
                {
                    AddLog("Komut Excel'e kaydedildi; görünüm yenilenemedi çünkü sayfa seçili değil.");
                    return;
                }

                try
                {
                    LoadSheet();
                    AddLog($"{payloads.Count} komut Excel'e kaydedildi ve görünüm yenilendi.");
                }
                catch (Exception ex)
                {
                    AddLog($"Komut Excel'e kaydedildi ancak görünüm yenilenemedi: {ex.Message}");
                }
            });
        }

        private void OnLogReceived(string message)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                AddLog(message);
            });
        }

        private void OnConnectionStatusChanged(bool isRunning)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                IsServerRunning = isRunning;
            });
        }

        private void SaveConfigFromExcel()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ExcelFilePath) || SelectedSheet == null)
                {
                    AddLog("Hata: Yapılandırmayı oluşturmak için Excel dosyası ve sayfası seçilmelidir.");
                    return;
                }

                _configService.SyncWithExcel(ExcelFilePath, SelectedSheet.Name);
                _configService.Save();
                AddLog("Yapılandırma Excel sayfasından başarıyla güncellendi ve kaydedildi.");
            }
            catch (Exception ex)
            {
                AddLog($"Yapılandırma kaydedilirken hata oluştu: {ex.Message}");
            }
        }
        private void SyncWithExcel()
        {
            try
            {
                if (string.IsNullOrEmpty(ExcelFilePath) || SelectedSheet == null)
                {
                    AddLog("Hata: Excel dosyası veya sayfa seçilmedi.");
                    return;
                }

                _configService.SyncWithExcel(ExcelFilePath, SelectedSheet.Name);
                _configService.Save();
                AddLog($"Şema başarıyla senkronize edildi: {SelectedSheet.Name} sayfasındaki yapılar kaydedildi.");
            }
            catch (Exception ex)
            {
                AddLog($"Şema senkronizasyonu sırasında hata: {ex.Message}");
            }
        }



    }

}
