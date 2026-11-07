using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using sttmerve_cc.Models;

namespace sttmerve_cc.Services
{
    public class HttpServerService
    {
        private const int MaxHeaderBytes = 16 * 1024;
        private const int MaxRequestBodyBytes = 1024 * 1024;

        private readonly int _port;
        private readonly ConfigService _configService;
        private readonly CommandRepository _repository;
        private readonly ExcelService _excelService;
        private readonly SemaphoreSlim _excelWriteLock = new(1, 1);
        private TcpListener? _listener;
        private volatile bool _isRunning;

        public event Action<List<CommandPayload>>? CommandReceived;
        public event Action<string>? LogMessageReceived;
        public event Action<bool>? ConnectionStatusChanged;
        public event Action<string>? ServerStarted;

        public HttpServerService(
            int port,
            ConfigService configService,
            CommandRepository repository,
            ExcelService excelService)
        {
            _port = port;
            _configService = configService;
            _repository = repository;
            _excelService = excelService;
        }

        public void Start()
        {
            if (_isRunning) return;

            var localIp = GetLocalIPAddress();
            var address = IPAddress.Parse(localIp);
            var listener = new TcpListener(address, _port);
            var url = $"http://{localIp}:{_port}/";

            try
            {
                listener.Start();
                _listener = listener;
                _isRunning = true;
                ConnectionStatusChanged?.Invoke(true);
                ServerStarted?.Invoke(url);
                LogMessageReceived?.Invoke($"Sunucu başlatıldı: {url}");
                _ = AcceptClientsAsync(listener);
            }
            catch (Exception ex)
            {
                listener.Stop();
                _listener = null;
                _isRunning = false;
                LogMessageReceived?.Invoke($"Sunucu başlatılamadı: {ex.Message}");
                throw;
            }
        }

        public void Stop()
        {
            if (!_isRunning) return;

            _isRunning = false;
            _listener?.Stop();
            _listener = null;
            ConnectionStatusChanged?.Invoke(false);
            LogMessageReceived?.Invoke("Sunucu durduruldu.");
        }

        private async Task AcceptClientsAsync(TcpListener listener)
        {
            while (_isRunning && ReferenceEquals(_listener, listener))
            {
                try
                {
                    var client = await listener.AcceptTcpClientAsync();
                    _ = ProcessClientAsync(client);
                }
                catch (SocketException ex)
                {
                    if (_isRunning && ReferenceEquals(_listener, listener))
                    {
                        LogMessageReceived?.Invoke($"Dinleyici hatası: {ex.Message}");
                        continue;
                    }

                    LogMessageReceived?.Invoke($"Dinleyici durduruldu: {ex.Message}");
                    break;
                }
                catch (ObjectDisposedException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    LogMessageReceived?.Invoke($"Dinleyici hatası: {ex.Message}");
                }
            }
        }

        private async Task ProcessClientAsync(TcpClient client)
        {
            using (client)
            {
                var stream = client.GetStream();
                try
                {
                    var request = await ReadRequestAsync(stream);
                    if (request == null) return;

                    LogMessageReceived?.Invoke($"İstek alındı: {request.Method} {request.Path}");
                    if (request.Path == "/config" && request.Method == "GET")
                    {
                        await HandleConfigRequestAsync(stream);
                    }
                    else if (request.Path == "/command" && request.Method == "POST")
                    {
                        await HandleCommandRequestAsync(stream, request.Body);
                    }
                    else
                    {
                        await SendErrorResponseAsync(
                            stream,
                            HttpStatusCode.NotFound,
                            "The requested endpoint was not found.");
                    }
                }
                catch (HttpProtocolException ex)
                {
                    await SendErrorResponseAsync(stream, ex.StatusCode, ex.Message);
                }
                catch (Exception ex)
                {
                    LogMessageReceived?.Invoke($"İstek işlenirken hata oluştu: {ex.Message}");
                    try
                    {
                        await SendErrorResponseAsync(
                            stream,
                            HttpStatusCode.InternalServerError,
                            "The request could not be completed.");
                    }
                    catch (IOException responseException)
                    {
                        LogMessageReceived?.Invoke(
                            $"Hata yanıtı gönderilemedi: {responseException.Message}");
                    }
                }
            }
        }

        private async Task HandleConfigRequestAsync(Stream stream)
        {
            var config = _configService.GetConfig();
            var json = JsonConvert.SerializeObject(config);
            await SendResponseAsync(stream, HttpStatusCode.OK, json);
        }

        private async Task HandleCommandRequestAsync(Stream stream, string json)
        {
            List<CommandPayload>? payloads;
            try
            {
                payloads = JsonConvert.DeserializeObject<List<CommandPayload>>(json);
            }
            catch (JsonException ex)
            {
                LogMessageReceived?.Invoke($"Komut JSON olarak çözümlenemedi: {ex.Message}");
                await SendErrorResponseAsync(
                    stream,
                    HttpStatusCode.BadRequest,
                    "The request body must be a JSON array of commands.");
                return;
            }

            if (payloads == null || payloads.Count == 0)
            {
                await SendErrorResponseAsync(
                    stream,
                    HttpStatusCode.BadRequest,
                    "The request must contain a non-empty JSON array of commands.");
                return;
            }

            if (payloads.Any(payload =>
                    payload == null ||
                    payload.Category == null ||
                    string.IsNullOrWhiteSpace(payload.Type) ||
                    string.IsNullOrWhiteSpace(payload.Value) ||
                    string.IsNullOrWhiteSpace(payload.BatchId)))
            {
                await SendErrorResponseAsync(
                    stream,
                    HttpStatusCode.BadRequest,
                    "Each command must include a batch ID, configured type, and non-empty value.");
                return;
            }

            var batchId = payloads[0].BatchId.Trim();
            if (payloads.Any(payload =>
                    !string.Equals(
                        payload.BatchId.Trim(),
                        batchId,
                        StringComparison.Ordinal)))
            {
                await SendErrorResponseAsync(
                    stream,
                    HttpStatusCode.BadRequest,
                    "A request may contain commands from only one batch.");
                return;
            }

            var config = _configService.GetConfig();
            if (string.IsNullOrWhiteSpace(config.ExcelFilePath) ||
                string.IsNullOrWhiteSpace(config.SheetName))
            {
                await SendErrorResponseAsync(
                    stream,
                    HttpStatusCode.Conflict,
                    "Synchronize the Excel sheet configuration before receiving commands.");
                return;
            }

            foreach (var payload in payloads)
            {
                var category = config.Categories.FirstOrDefault(item =>
                    string.Equals(
                        item.Name.Trim(),
                        payload.Category.Trim(),
                        StringComparison.OrdinalIgnoreCase));
                if (category == null ||
                    !category.Types.Any(type =>
                        string.Equals(
                            type.Trim(),
                            payload.Type.Trim(),
                            StringComparison.OrdinalIgnoreCase)))
                {
                    await SendErrorResponseAsync(
                        stream,
                        HttpStatusCode.BadRequest,
                        $"Unconfigured category/type pair: '{payload.Category}' / '{payload.Type}'.");
                    return;
                }
            }

            var alreadyProcessed = false;
            var excelRow = 0;
            await _excelWriteLock.WaitAsync();
            try
            {
                alreadyProcessed = _repository.HasProcessedBatch(batchId);
                if (!alreadyProcessed)
                {
                    excelRow = _excelService.WriteCommandBatchToExcel(
                        config.SheetName,
                        payloads,
                        batchId);
                    _repository.SaveCommandBatch(batchId, payloads);
                }
            }
            finally
            {
                _excelWriteLock.Release();
            }

            if (!alreadyProcessed)
            {
                CommandReceived?.Invoke(payloads);
                LogMessageReceived?.Invoke(
                    $"Komut grubu Excel'e kaydedildi. Batch: {batchId}, satır: {excelRow}.");
            }
            else
            {
                LogMessageReceived?.Invoke($"Yinelenen komut grubu atlandı: {batchId}.");
            }

            await SendSuccessResponseAsync(stream, batchId, excelRow);
        }

        private static async Task<HttpRequestData?> ReadRequestAsync(Stream stream)
        {
            using var headerBuffer = new MemoryStream();
            var nextByte = new byte[1];
            var hasTerminator = false;
            while (headerBuffer.Length < MaxHeaderBytes)
            {
                var count = await stream.ReadAsync(nextByte, 0, 1);
                if (count == 0)
                {
                    if (headerBuffer.Length == 0) return null;
                    throw new HttpProtocolException(
                        HttpStatusCode.BadRequest,
                        "The HTTP request headers are incomplete.");
                }

                headerBuffer.WriteByte(nextByte[0]);
                var length = headerBuffer.Length;
                if (length >= 4)
                {
                    var bytes = headerBuffer.GetBuffer();
                    hasTerminator =
                        bytes[length - 4] == 13 &&
                        bytes[length - 3] == 10 &&
                        bytes[length - 2] == 13 &&
                        bytes[length - 1] == 10;
                    if (hasTerminator) break;
                }
            }

            if (!hasTerminator)
            {
                throw new HttpProtocolException(
                    HttpStatusCode.BadRequest,
                    "The HTTP request headers are too large.");
            }

            var headerText = Encoding.ASCII.GetString(
                headerBuffer.GetBuffer(),
                0,
                (int)headerBuffer.Length);
            var lines = headerText.Split(new[] { "\r\n" }, StringSplitOptions.None);
            var requestParts = lines[0].Split(' ');
            if (requestParts.Length != 3 || !requestParts[2].StartsWith("HTTP/"))
            {
                throw new HttpProtocolException(
                    HttpStatusCode.BadRequest,
                    "The HTTP request line is invalid.");
            }

            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 1; index < lines.Length - 2; index++)
            {
                var separator = lines[index].IndexOf(':');
                if (separator <= 0)
                {
                    throw new HttpProtocolException(
                        HttpStatusCode.BadRequest,
                        "An HTTP request header is invalid.");
                }

                var name = lines[index].Substring(0, separator).Trim();
                var value = lines[index].Substring(separator + 1).Trim();
                if (!headers.TryAdd(name, value))
                {
                    throw new HttpProtocolException(
                        HttpStatusCode.BadRequest,
                        "An HTTP request header is duplicated.");
                }
            }

            if (headers.TryGetValue("Transfer-Encoding", out var transferEncoding) &&
                !string.Equals(transferEncoding, "identity", StringComparison.OrdinalIgnoreCase))
            {
                throw new HttpProtocolException(
                    HttpStatusCode.BadRequest,
                    "Transfer-encoded request bodies are not supported.");
            }

            var contentLength = 0;
            if (headers.TryGetValue("Content-Length", out var contentLengthText))
            {
                if (!int.TryParse(
                        contentLengthText,
                        NumberStyles.None,
                        CultureInfo.InvariantCulture,
                        out contentLength) ||
                    contentLength < 0)
                {
                    throw new HttpProtocolException(
                        HttpStatusCode.BadRequest,
                        "The HTTP content length is invalid.");
                }
            }
            else if (string.Equals(requestParts[0], "POST", StringComparison.OrdinalIgnoreCase))
            {
                throw new HttpProtocolException(
                    HttpStatusCode.LengthRequired,
                    "A content length is required.");
            }

            if (contentLength > MaxRequestBodyBytes)
            {
                throw new HttpProtocolException(
                    HttpStatusCode.RequestEntityTooLarge,
                    "The request body is too large.");
            }

            var bodyBytes = new byte[contentLength];
            var bodyOffset = 0;
            while (bodyOffset < bodyBytes.Length)
            {
                var count = await stream.ReadAsync(
                    bodyBytes,
                    bodyOffset,
                    bodyBytes.Length - bodyOffset);
                if (count == 0)
                {
                    throw new HttpProtocolException(
                        HttpStatusCode.BadRequest,
                        "The HTTP request body is incomplete.");
                }

                bodyOffset += count;
            }

            var target = requestParts[1];
            var path = Uri.TryCreate(target, UriKind.Absolute, out var absoluteUri)
                ? absoluteUri.AbsolutePath
                : target.Split('?')[0];
            return new HttpRequestData(
                requestParts[0].ToUpperInvariant(),
                path,
                Encoding.UTF8.GetString(bodyBytes));
        }

        private static Task SendSuccessResponseAsync(
            Stream stream,
            string batchId,
            int excelRow)
        {
            return SendResponseAsync(
                stream,
                HttpStatusCode.OK,
                JsonConvert.SerializeObject(new
                {
                    status = "success",
                    batchId,
                    excelRow
                }));
        }

        private static Task SendErrorResponseAsync(
            Stream stream,
            HttpStatusCode statusCode,
            string message)
        {
            return SendResponseAsync(
                stream,
                statusCode,
                JsonConvert.SerializeObject(new { status = "error", message }));
        }

        private static async Task SendResponseAsync(
            Stream stream,
            HttpStatusCode statusCode,
            string body)
        {
            var bodyBytes = Encoding.UTF8.GetBytes(body);
            var reason = GetReasonPhrase(statusCode);
            var headers = Encoding.ASCII.GetBytes(
                $"HTTP/1.1 {(int)statusCode} {reason}\r\n" +
                "Content-Type: application/json; charset=utf-8\r\n" +
                $"Content-Length: {bodyBytes.Length}\r\n" +
                "Connection: close\r\n\r\n");
            await stream.WriteAsync(headers, 0, headers.Length);
            await stream.WriteAsync(bodyBytes, 0, bodyBytes.Length);
            await stream.FlushAsync();
        }

        private static string GetReasonPhrase(HttpStatusCode statusCode)
        {
            return statusCode switch
            {
                HttpStatusCode.OK => "OK",
                HttpStatusCode.BadRequest => "Bad Request",
                HttpStatusCode.Conflict => "Conflict",
                HttpStatusCode.LengthRequired => "Length Required",
                HttpStatusCode.NotFound => "Not Found",
                HttpStatusCode.RequestEntityTooLarge => "Payload Too Large",
                HttpStatusCode.InternalServerError => "Internal Server Error",
                _ => statusCode.ToString()
            };
        }

        private string GetLocalIPAddress()
        {
            try
            {
                foreach (var networkInterface in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (networkInterface.OperationalStatus != OperationalStatus.Up ||
                        (networkInterface.NetworkInterfaceType != NetworkInterfaceType.Wireless80211 &&
                         networkInterface.NetworkInterfaceType != NetworkInterfaceType.Ethernet))
                    {
                        continue;
                    }

                    foreach (var address in networkInterface.GetIPProperties().UnicastAddresses)
                    {
                        if (address.Address.AddressFamily == AddressFamily.InterNetwork &&
                            !IPAddress.IsLoopback(address.Address))
                        {
                            return address.Address.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessageReceived?.Invoke($"Ağ adresi alınamadı: {ex.Message}");
            }

            return IPAddress.Loopback.ToString();
        }

        private sealed class HttpRequestData
        {
            public string Method { get; }
            public string Path { get; }
            public string Body { get; }

            public HttpRequestData(string method, string path, string body)
            {
                Method = method;
                Path = path;
                Body = body;
            }
        }

        private sealed class HttpProtocolException : Exception
        {
            public HttpStatusCode StatusCode { get; }

            public HttpProtocolException(HttpStatusCode statusCode, string message)
                : base(message)
            {
                StatusCode = statusCode;
            }
        }
    }
}
