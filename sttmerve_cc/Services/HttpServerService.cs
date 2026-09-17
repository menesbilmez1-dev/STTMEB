using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using sttmerve_cc.Models;
using sttmerve_cc.Services;

namespace sttmerve_cc.Services
{
    public class HttpServerService
    {
        private readonly HttpListener _listener;
        private readonly int _port;
        private readonly ConfigService _configService;
        private bool _isRunning;

        public event Action<List<CommandPayload>>? CommandReceived;
        public event Action<string>? LogMessageReceived;
        public event Action<bool>? ConnectionStatusChanged;
        public event Action<string>? ServerStarted;
        private readonly CommandRepository _repository;
     
        private readonly ExcelService _excelService;   // <--- ADD THIS (to sync to Excel)

        public HttpServerService(int port, ConfigService configService, CommandRepository repository, ExcelService excelService)
        {
            _port = port;
            _configService = configService;
            _repository = repository;
            _listener = new HttpListener();
            _excelService = excelService;
        }

        public void Start()
        {
            if (_isRunning) return;

            string localIp = GetLocalIPAddress();
            string prefix = $"http://{localIp}:{_port}/";
           // string localhostPrefix = $"http://localhost:{_port}/";

            try
            {
                _listener.Prefixes.Add(prefix);
             

                _listener.Start();
                _isRunning = true;
                ConnectionStatusChanged?.Invoke(true);
                ServerStarted?.Invoke(prefix);
                LogMessageReceived?.Invoke($"Sunucu başlatıldı: {prefix}");
            }
            catch (Exception ex)
            {
                LogMessageReceived?.Invoke($"Sunucu başlatılamadı: {ex.Message}");
                throw;
            }

            Task.Run(async () =>
            {
                while (_isRunning)
                {
                    try
                    {
                        var context = await _listener.GetContextAsync();
                        Console.WriteLine($"Alındı: {context.Request.HttpMethod} {context.Request.Url.AbsolutePath}");
                        _ = ProcessRequest(context);
                    }
                    catch (Exception ex) when (_isRunning)
                    {
                        LogMessageReceived?.Invoke($"Dinleyici hatası: {ex.Message}");
                    }
                }
            });
        }

        public void Stop()
        {
            _isRunning = false;
            _listener.Stop();
            ConnectionStatusChanged?.Invoke(false);
            LogMessageReceived?.Invoke("Sunucu durduruldu.");
        }

        private async Task ProcessRequest(HttpListenerContext context)
        {
            try
            {
                string path = context.Request.Url?.AbsolutePath ?? "";
                LogMessageReceived?.Invoke($"İstek alındı: {context.Request.HttpMethod} {path}");

                if (path == "/config" && context.Request.HttpMethod == "GET")
                {
                    await HandleConfigRequest(context);
                }
                else if (path == "/command" && context.Request.HttpMethod == "POST")
                {
                    await HandleCommandRequest(context);
                }
                else
                {
                    context.Response.StatusCode = (int)HttpStatusCode.NotFound;
                    context.Response.Close();
                }
            }
            catch (Exception ex)
            {
                LogMessageReceived?.Invoke($"Error processing request: {ex.Message}");
                try
                {
                    context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                    context.Response.Close();
                }
                catch
                {
                    // Ignore errors during cleanup
                }
            }
        }

        private async Task HandleConfigRequest(HttpListenerContext context)
        {
            try
            {
                var config = _configService.GetConfig();
                string json = JsonConvert.SerializeObject(config);
                byte[] buffer = Encoding.UTF8.GetBytes(json);

                context.Response.ContentType = "application/json";
                context.Response.StatusCode = (int)HttpStatusCode.OK;
                await context.Response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                context.Response.Close();
            }
            catch (Exception ex)
            {
                LogMessageReceived?.Invoke($"Error in HandleConfigRequest: {ex.Message}");
                context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                context.Response.Close();
            }
        }
        private async Task ProcessPayload(CommandPayload payload)
        {
            
            _repository.SaveCommand(payload);

            
          

            Console.WriteLine($"[HTTP] Saved and Exported: {payload.Type}");
        }
        private async Task HandleCommandRequest(HttpListenerContext context)
        {
            try
            {
                using var reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding);
                string json = await reader.ReadToEndAsync();

                LogMessageReceived?.Invoke($"Alındı: {json}");

                var payloads = JsonConvert.DeserializeObject<List<CommandPayload>>(json);
                if (payloads != null && payloads.Count > 0)
                {
                    CommandReceived?.Invoke(payloads);
                    foreach (var item in payloads)
                    {
                        _repository.SaveCommand(item);
                    }
                    await SendSuccessResponse(context);
                }
                else
                {
                    var singlePayload = JsonConvert.DeserializeObject<CommandPayload>(json);
                    if (singlePayload != null)
                    {
                        CommandReceived?.Invoke(new List<CommandPayload> { singlePayload });
                        _repository.SaveCommand(singlePayload);
                        await SendSuccessResponse(context);
                    }
                    else
                    {
                        context.Response.StatusCode = (int)HttpStatusCode.BadRequest;
                        context.Response.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessageReceived?.Invoke($"HandleCommandRequest sırasında hata: {ex.Message}");
                context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                context.Response.Close();
            }
        }

        private async Task SendSuccessResponse(HttpListenerContext context)
        {
            try
            {
                context.Response.StatusCode = (int)HttpStatusCode.OK;
                byte[] buffer = Encoding.UTF8.GetBytes("{\"status\":\"success\"}");
                await context.Response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                context.Response.Close();
            }
            catch (Exception ex)
            {
                LogMessageReceived?.Invoke($"Başarı yanıtı gönderilirken hata oluştu: {ex.Message}");
            }
        }

        private string GetLocalIPAddress()
        {
            try
            {
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (ni.OperationalStatus == OperationalStatus.Up &&
                        (ni.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 ||
                         ni.NetworkInterfaceType == NetworkInterfaceType.Ethernet))
                    {
                        foreach (UnicastIPAddressInformation ip in ni.GetIPProperties().UnicastAddresses)
                        {
                            if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                            {
                                return ip.Address.ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Fallback to localhost if something goes wrong during discovery
            }
            return "127.0.0.1";
        }
    }
}
