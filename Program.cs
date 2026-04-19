using System.Text.Json;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace AlexaToExcel
{
    class Program
    {
        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        private static extern bool FreeConsole();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

        private static bool _consoleAllocated;

        private static void EnsureConsole()
        {
            if (!_consoleAllocated)
            {
                AllocConsole();
                _consoleAllocated = true;
            }
        }

        private static void ReleaseConsole()
        {
            if (_consoleAllocated)
            {
                FreeConsole();
                _consoleAllocated = false;
            }
        }

        static async Task Main(string[] args)
        {
            bool needsConsole = args.Contains("--debug", StringComparer.OrdinalIgnoreCase) ||
                                args.Contains("--help", StringComparer.OrdinalIgnoreCase) ||
                                args.Contains("-h", StringComparer.OrdinalIgnoreCase) ||
                                Array.Exists(args, a => a.Equals("--html-to-csv", StringComparison.OrdinalIgnoreCase));

            if (needsConsole)
            {
                EnsureConsole();
            }
            else
            {
                // Detach from any inherited console (e.g. launched from PowerShell/cmd)
                // so no window is visible during normal background operation.
                FreeConsole();
            }

            Console.WriteLine("=== Alexa To Excel ===");
            Console.WriteLine();

            if (args.Contains("--help", StringComparer.OrdinalIgnoreCase) ||
                args.Contains("-h", StringComparer.OrdinalIgnoreCase))
            {
                PrintHelp();
                return;
            }

            // --html-to-csv <input.html> [output.csv]
            int htmlIdx = Array.FindIndex(args, a => a.Equals("--html-to-csv", StringComparison.OrdinalIgnoreCase));
            if (htmlIdx >= 0)
            {
                RunHtmlToCsv(args, htmlIdx);
                return;
            }

            var config = AppConfig.Load();

            // --debug flag: dump raw API response and exit
            bool debug = args.Contains("--debug", StringComparer.OrdinalIgnoreCase);

            ApplyArgs(args, config);

            Console.WriteLine($"Output file  : {config.OutputPath}");
            Console.WriteLine($"Poll interval: {config.PollIntervalMinutes} min");
            Console.WriteLine($"Alexa host   : {AlexaReminderService.GetAlexaHost(config.BaseUrl)}");
            Console.WriteLine();

            // Cookie is required – must be set in config.json
            if (string.IsNullOrWhiteSpace(config.CookieString))
            {
                if (!RunLoginFlow(config))
                {
                    return;
                }
            }

            var service = new AlexaReminderService(config, debug);

            try
            {
                //OpenBrowser(config);
                await RunSync(service, config);
            }
            catch (AuthException)
            {
                if (!RunLoginFlow(config))
                {
                    return;
                }
                service = new AlexaReminderService(config, debug);
                try
                {
                    await RunSync(service, config);
                }
                catch (AuthException ex2)
                {
                    MessageBoxW(IntPtr.Zero, $"Still getting 401 after re-login:\n{ex2.Message}", "Alexa Reminder Sync", 0x10);
                    return;
                }
            }

            if (!debug && config.PollIntervalMinutes > 0)
            {
                ReleaseConsole();
                while (true)
                {
                    await Task.Delay(TimeSpan.FromMinutes(config.PollIntervalMinutes));
                    try
                    {
                        await RunSync(service, config);
                    }
                    catch (AuthException)
                    {
                        if (!RunLoginFlow(config))
                        {
                            return;
                        }
                        service = new AlexaReminderService(config, debug);
                    }
                }
            }

            if (debug)
            {
                Console.WriteLine();
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }


        // Returns true if login succeeded and cookies were saved; false if the user cancelled.
        static bool RunLoginFlow(AppConfig config)
        {
            bool success = false;

            var staThread = new Thread(() =>
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                using var form = new LoginForm(config);
                var result = form.ShowDialog();

                if (result != DialogResult.OK || string.IsNullOrWhiteSpace(form.ExtractedCookie))
                {
                    MessageBoxW(IntPtr.Zero, "Login cancelled or cookies not detected. Exiting.", "Alexa Reminder Sync", 0x30);
                    success = false;
                }
                else
                {
                    config.CookieString = form.ExtractedCookie;
                    config.Save();
                    success = true;
                }
            });

            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();

            return success;
        }

        static async Task RunSync(AlexaReminderService service, AppConfig config)
        {
            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Fetching...");
            try
            {
                var reminders = await service.FetchRemindersAsync();
                Console.WriteLine($"  Fetched {reminders.Count} reminder(s).");

                var writer = new ExcelWriter(config.OutputPath);
                int added = writer.MergeAndSave(reminders);
                Console.WriteLine($"  {added} new row(s) written to {config.OutputPath}");
            }
            catch (AuthException ex)
            {
                Console.WriteLine($"  AUTH ERROR: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ERROR: {ex.Message}");
            }
        }

        static void RunHtmlToCsv(string[] args, int flagIndex)
        {
            if (flagIndex + 1 >= args.Length)
            {
                Console.WriteLine("Error: --html-to-csv requires an input HTML file path.");
                Console.WriteLine("Usage: AlexaToExcel --html-to-csv <input.html> [output.csv]");
                return;
            }

            string htmlPath = args[flagIndex + 1];
            string csvPath;

            if (flagIndex + 2 < args.Length && !args[flagIndex + 2].StartsWith("--"))
            {
                csvPath = args[flagIndex + 2];
            }
            else
            {
                csvPath = Path.ChangeExtension(htmlPath, ".csv");
            }

            Console.WriteLine($"Converting HTML to CSV...");
            Console.WriteLine($"  Input : {htmlPath}");
            Console.WriteLine($"  Output: {csvPath}");

            try
            {
                int rows = HtmlToCsvConverter.Convert(htmlPath, csvPath);
                Console.WriteLine($"  Done! {rows} row(s) written to {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ERROR: {ex.Message}");
            }
        }

        static void PrintHelp()
        {
            Console.WriteLine("Usage: AlexaToExcel [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  -c, --country <code>     Country code (overrides BaseUrl in config.json)");
            Console.WriteLine("                           Codes: us, ca, uk, de, fr, es, it, au, jp, in, mx, br");
            Console.WriteLine("  -p, --poll-interval <m>  Poll interval in minutes (overrides config.json).");
            Console.WriteLine("                           Use 0 to run once and exit.");
            Console.WriteLine("  --html-to-csv <file> [out.csv]");
            Console.WriteLine("                           Parse tables from an HTML file and export to CSV.");
            Console.WriteLine("                           If output path is omitted, uses the same name with .csv.");
            Console.WriteLine("  --debug                  Show detailed HTTP request/response info.");
            Console.WriteLine("  --help                   Show this help message.");
        }

        static void ApplyArgs(string[] args, AppConfig config)
        {
            for (int i = 0; i < args.Length; i++)
            {
                if (TryGetArgValue(args, ref i, out var country, "--country", "-c"))
                {
                    var url = CountryToBaseUrl(country);
                    if (url == null)
                    {
                        Console.WriteLine($"Unknown country code '{country}'. Valid codes: us, ca, uk, de, fr, es, it, au, jp, in, mx, br");
                        Console.WriteLine("Falling back to BaseUrl in config.json.");
                    }
                    else
                    {
                        config.BaseUrl = url;
                    }
                }
                else if (TryGetArgValue(args, ref i, out var intervalStr, "--poll-interval", "-p"))
                {
                    if (int.TryParse(intervalStr, out int interval) && interval >= 0)
                    {
                        config.PollIntervalMinutes = interval;
                    }
                    else
                    {
                        Console.WriteLine($"Invalid poll interval '{intervalStr}'. Must be a non-negative integer.");
                    }
                }
            }
        }

        static bool TryGetArgValue(string[] args, ref int i, out string value, params string[] flags)
        {
            var arg = args[i];
            foreach (var flag in flags)
            {
                var prefix = flag + "=";
                if (arg.Equals(flag, StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    value = args[++i];
                    return true;
                }
                if (arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    value = arg[prefix.Length..];
                    return true;
                }
            }
            value = "";
            return false;
        }

        static string? CountryToBaseUrl(string code) => code.ToLowerInvariant() switch
        {
            "us" or "usa"           => "https://www.amazon.com",
            "ca" or "canada"        => "https://www.amazon.ca",
            "uk" or "gb"            => "https://www.amazon.co.uk",
            "de" or "germany"       => "https://www.amazon.de",
            "fr" or "france"        => "https://www.amazon.fr",
            "es" or "spain"         => "https://www.amazon.es",
            "it" or "italy"         => "https://www.amazon.it",
            "au" or "australia"     => "https://www.amazon.com.au",
            "jp" or "japan"         => "https://www.amazon.co.jp",
            "in" or "india"         => "https://www.amazon.in",
            "mx" or "mexico"        => "https://www.amazon.com.mx",
            "br" or "brazil"        => "https://www.amazon.com.br",
            _ => null
        };
    }

    // ─── Exceptions ────────────────────────────────────────────────────────────

    class AuthException : Exception
    {
        public AuthException(string msg) : base(msg) { }
    }

    // ─── Config ────────────────────────────────────────────────────────────────

    class AppConfig
    {
        public string BaseUrl { get; set; } = "https://www.amazon.com";
        public string CookieString { get; set; } = "";
        public string OutputPath { get; set; } = "alexa_reminders.xlsx";
        public int PollIntervalMinutes { get; set; } = 60;

        private static readonly string ConfigFile = "config.json";

        public static AppConfig Load()
        {
            if (File.Exists(ConfigFile))
            {
                try
                {
                    var json = File.ReadAllText(ConfigFile);
                    return JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: could not parse config.json: {ex.Message}");
                }
            }
            var cfg = new AppConfig();
            cfg.Save();
            return cfg;
        }

        public void Save()
        {
            var opts = new JsonSerializerOptions { WriteIndented = true };
            File.WriteAllText(ConfigFile, JsonSerializer.Serialize(this, opts));
        }
    }

    // ─── Models ────────────────────────────────────────────────────────────────

    class AlexaReminder
    {
        public string Id { get; set; } = "";
        public string Text { get; set; } = "";
        public DateTime? TriggerTime { get; set; }
        public string Status { get; set; } = "";
        public string RecurrenceType { get; set; } = "";
        public string CreatedDate { get; set; } = "";
        public string DeviceName { get; set; } = "";
        public string RawJson { get; set; } = "";
    }

    record AlexaDevice(string SerialNumber, string DeviceType, string FriendlyName);

    // ─── Service ───────────────────────────────────────────────────────────────

    class AlexaReminderService
    {
        private readonly AppConfig _config;
        private readonly HttpClient _http;
        private readonly bool _debug;
        private string _alexaHost;

        public static string GetAlexaHost(string baseUrl)
        {
            var url = (baseUrl ?? "https://www.amazon.com").Trim().TrimEnd('/');
            if (url.Contains("alexa.amazon.")) return url;
            if (url.Contains("://www.amazon."))  return url.Replace("://www.amazon.",  "://alexa.amazon.");
            if (url.Contains("://amazon."))       return url.Replace("://amazon.",       "://alexa.amazon.");
            return url; // fallback – shouldn't happen
        }

        public AlexaReminderService(AppConfig config, bool debug = false)
        {
            _config = config;
            _debug  = debug;
            _alexaHost = GetAlexaHost(config.BaseUrl);

            var cookie = CleanCookie(config.CookieString);
            var csrf   = ExtractCsrf(cookie);

            if (_debug)
            {
                Console.WriteLine($"[DEBUG] AlexaHost   : {_alexaHost}");
                Console.WriteLine($"[DEBUG] Cookie len  : {cookie.Length} chars");
                Console.WriteLine($"[DEBUG] csrf value  : {(csrf ?? "(NOT FOUND)")}");
                Console.WriteLine($"[DEBUG] Cookie names: {string.Join(", ", cookie.Split(';').Select(p => p.Trim().Split('=')[0]).Where(n => n.Length > 0))}");
            }

            if (csrf == null)
            {
                Console.WriteLine("WARNING: 'csrf' cookie not found in cookie string.");
                Console.WriteLine("         API calls will likely fail. See README for help.");
            }
            else
            {
                Console.WriteLine($"  csrf  : {csrf}");
            }

            var handler = new HttpClientHandler
            {
                UseCookies = false,
                AllowAutoRedirect = false,
                AutomaticDecompression = System.Net.DecompressionMethods.All
            };
            _http = new HttpClient(handler);
            _http.Timeout = TimeSpan.FromSeconds(30);

            // Match exactly what Chrome sends
            _http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
                "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("Accept",
                "application/json, text/javascript, */*; q=0.01");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("Accept-Language",
                "en-CA,en;q=0.9");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("Accept-Encoding",
                "gzip, deflate, br");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("Cache-Control", "no-cache");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("Pragma", "no-cache");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("Origin", _alexaHost);
            _http.DefaultRequestHeaders.TryAddWithoutValidation("Referer",
                $"{_alexaHost}/spa/index.html#reminders");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("sec-ch-ua",
                "\"Google Chrome\";v=\"124\", \"Chromium\";v=\"124\", \"Not-A.Brand\";v=\"99\"");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("sec-ch-ua-mobile", "?0");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("sec-ch-ua-platform", "\"Windows\"");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("sec-fetch-dest", "empty");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("sec-fetch-mode", "cors");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("sec-fetch-site", "same-origin");
            _http.DefaultRequestHeaders.TryAddWithoutValidation("Cookie", cookie);
            _http.DefaultRequestHeaders.TryAddWithoutValidation("X-Requested-With",
                "XMLHttpRequest");

            if (csrf != null)
                _http.DefaultRequestHeaders.TryAddWithoutValidation("csrf", csrf);
        }

        // Clean up common paste artifacts from config.json / terminal
        static string CleanCookie(string raw)
        {
            var s = raw.Trim();
            // Strip surrounding double-quotes (JSON copy artifact)
            if (s.StartsWith("\"") && s.EndsWith("\"") && s.Length > 1)
                s = s[1..^1];
            return s;
        }

        static string? ExtractCsrf(string cookie)
        {
            foreach (var part in cookie.Split(';'))
            {
                var p = part.Trim();
                if (p.StartsWith("csrf=", StringComparison.OrdinalIgnoreCase))
                {
                    var val = p[5..].Trim();
                    // Strip Windows cmd escape chars and stray quotes
                    val = val.TrimEnd('^', '%').Trim('"', '\'');
                    return val.Length > 0 ? val : null;
                }
            }
            return null;
        }

        public async Task<List<AlexaReminder>> FetchRemindersAsync()
        {
            // ── Step 1: get devices ─────────────────────────────────────────
            var devUrl = $"{_alexaHost}/api/devices-v2/device?raw=false";
            if (_debug) Console.WriteLine($"[DEBUG] GET {devUrl}");

            var devResp = await SendGet(devUrl);
            var devJson = await devResp.Content.ReadAsStringAsync();

            if (_debug)
            {
                Console.WriteLine($"[DEBUG] Devices response ({devJson.Length} chars):");
                Console.WriteLine(devJson[..Math.Min(800, devJson.Length)]);
            }

            var devices = ParseDevices(devJson);
            Console.WriteLine($"  Devices: {devices.Count} found" +
                (devices.Count > 0 ? $" ({string.Join(", ", devices.Select(d => d.FriendlyName))})" : ""));

            if (devices.Count == 0)
                throw new Exception("No Alexa devices found. The API returned an empty device list.");

            // ── Step 2: get notifications per device ────────────────────────
            var all    = new List<AlexaReminder>();
            var seen   = new HashSet<string>();

            foreach (var dev in devices)
            {
                var url = $"{_alexaHost}/api/notifications" +
                          $"?deviceSerialNumber={Uri.EscapeDataString(dev.SerialNumber)}" +
                          $"&deviceType={Uri.EscapeDataString(dev.DeviceType)}";

                if (_debug) Console.WriteLine($"[DEBUG] GET {url}");

                try
                {
                    var resp = await SendGet(url);
                    var body = await resp.Content.ReadAsStringAsync();

                    if (_debug)
                    {
                        Console.WriteLine($"[DEBUG] Notifications for {dev.FriendlyName} ({body.Length} chars):");
                        Console.WriteLine(body[..Math.Min(600, body.Length)]);
                    }

                    var items = ParseNotifications(body, dev.FriendlyName);
                    int before = all.Count;
                    foreach (var r in items)
                    {
                        var key = r.Id + "|" + r.Text;
                        if (seen.Add(key)) all.Add(r);
                    }
                    Console.WriteLine($"    {dev.FriendlyName}: {all.Count - before} reminder(s)");
                }
                catch (AuthException) { throw; }
                catch (Exception ex)
                {
                    Console.WriteLine($"    {dev.FriendlyName}: skipped ({ex.Message})");
                }
            }

            return all.OrderByDescending(r => r.TriggerTime ?? DateTime.MinValue).ToList();
        }

        async Task<HttpResponseMessage> SendGet(string url)
        {
            var resp = await _http.GetAsync(url);

            if (_debug)
                Console.WriteLine($"[DEBUG] HTTP {(int)resp.StatusCode} {resp.ReasonPhrase}");

            if (resp.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                resp.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                var body = await resp.Content.ReadAsStringAsync();
                throw new AuthException(
                    $"HTTP {(int)resp.StatusCode} from {url}\n" +
                    $"  Server said: {body[..Math.Min(300, body.Length)]}");
            }

            if (!resp.IsSuccessStatusCode)
            {
                var body = await resp.Content.ReadAsStringAsync();
                throw new Exception(
                    $"HTTP {(int)resp.StatusCode} from {url}: {body[..Math.Min(300, body.Length)]}");
            }
            return resp;
        }

        List<AlexaDevice> ParseDevices(string json)
        {
            var list = new List<AlexaDevice>();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (!root.TryGetProperty("devices", out var arr)) return list;

            foreach (var item in arr.EnumerateArray())
            {
                var serial = Str(item, "serialNumber", "deviceSerialNumber");
                var type   = Str(item, "deviceType");
                var name   = Str(item, "accountName", "name") ?? serial ?? "Unknown";
                if (serial != null && type != null)
                    list.Add(new AlexaDevice(serial, type, name));
            }
            return list;
        }

        List<AlexaReminder> ParseNotifications(string json, string deviceName)
        {
            var list = new List<AlexaReminder>();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (!root.TryGetProperty("notifications", out var arr)) return list;

            foreach (var item in arr.EnumerateArray())
            {
                var type = Str(item, "type") ?? "";
                // Keep REMINDER type; also keep items with no type (some regions omit it)
                if (type.Length > 0 && !type.Contains("REMINDER", StringComparison.OrdinalIgnoreCase))
                    continue;

                var r = new AlexaReminder
                {
                    Id         = Str(item, "notificationIndex", "id", "notificationId") ?? Guid.NewGuid().ToString(),
                    Status     = Str(item, "status") ?? "",
                    DeviceName = deviceName,
                    Text       = Str(item, "reminderLabel", "alarmLabel", "text", "value") ?? "",
                    CreatedDate = Str(item, "createdDate", "createdTime") ?? "",
                    RawJson    = item.GetRawText()
                };

                // Time: combine originalDate + originalTime (Alexa's standard split fields)
                var origDate = Str(item, "originalDate");
                var origTime = Str(item, "originalTime");
                if (origDate != null && origTime != null)
                {
                    ApplyTime(origDate + "T" + origTime.TrimStart('T'), r);
                }

                // Fallback to unified timestamp / datetime fields
                if (!r.TriggerTime.HasValue)
                {
                    ApplyTime(Str(item, "triggerTime", "alarmTime", "scheduledTime"), r);
                }

                // Time from nested trigger object
                if (!r.TriggerTime.HasValue && item.TryGetProperty("trigger", out var trig))
                {
                    ApplyTime(Str(trig, "scheduledTime", "originalTime", "alarmTime"), r);
                    r.RecurrenceType = Str(trig, "type", "recurrenceRule") ?? "";
                }

                list.Add(r);
            }
            return list;
        }

        static void ApplyTime(string? val, AlexaReminder r)
        {
            if (string.IsNullOrEmpty(val)) return;
            if (long.TryParse(val, out long n))
            {
                r.TriggerTime = (n > 1_000_000_000_000L)
                    ? DateTimeOffset.FromUnixTimeMilliseconds(n).LocalDateTime
                    : DateTimeOffset.FromUnixTimeSeconds(n).LocalDateTime;
            }
            else if (DateTime.TryParse(val, out var dt))
            {
                r.TriggerTime = dt;
            }
        }

        static string? Str(JsonElement el, params string[] keys)
        {
            foreach (var k in keys)
                if (el.TryGetProperty(k, out var v))
                {
                    if (v.ValueKind == JsonValueKind.String) return v.GetString();
                    if (v.ValueKind == JsonValueKind.Number) return v.GetRawText();
                }
            return null;
        }
    }

    // ─── Excel ─────────────────────────────────────────────────────────────────

    class ExcelWriter
    {
        private readonly string _path;
        public ExcelWriter(string path) => _path = path;

        public int MergeAndSave(List<AlexaReminder> incoming)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var existingIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            ExcelPackage pkg;

            if (File.Exists(_path))
            {
                pkg = new ExcelPackage(new FileInfo(_path));
                var ws = pkg.Workbook.Worksheets["Reminders"];
                if (ws?.Dimension != null)
                    for (int row = 2; row <= ws.Dimension.End.Row; row++)
                    {
                        var id = ws.Cells[row, 1].Text;
                        if (id.Length > 0) existingIds.Add(id);
                    }
            }
            else
            {
                pkg = new ExcelPackage();
            }

            var sheet = pkg.Workbook.Worksheets["Reminders"]
                     ?? pkg.Workbook.Worksheets.Add("Reminders");

            if (sheet.Dimension == null) WriteHeader(sheet);

            int nextRow = (sheet.Dimension?.End.Row ?? 1) + 1;
            int added = 0;

            foreach (var r in incoming)
            {
                if (existingIds.Contains(r.Id)) continue;

                sheet.Cells[nextRow, 1].Value = r.Id;
                sheet.Cells[nextRow, 2].Value = r.Text;
                sheet.Cells[nextRow, 3].Value = r.TriggerTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "";
                sheet.Cells[nextRow, 4].Value = r.Status;
                sheet.Cells[nextRow, 5].Value = r.RecurrenceType;
                sheet.Cells[nextRow, 6].Value = r.CreatedDate;
                sheet.Cells[nextRow, 7].Value = r.DeviceName;
                sheet.Cells[nextRow, 8].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                if (nextRow % 2 == 0)
                {
                    sheet.Cells[nextRow, 1, nextRow, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[nextRow, 1, nextRow, 8].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
                }

                nextRow++;
                added++;
            }

            if (added > 0)
                for (int col = 1; col <= 8; col++)
                    sheet.Column(col).AutoFit();

            pkg.SaveAs(new FileInfo(_path));
            return added;
        }

        void WriteHeader(ExcelWorksheet ws)
        {
            string[] headers = { "ID", "Reminder Text", "Trigger Time", "Status",
                                  "Recurrence", "Created Date", "Device", "Synced At" };
            for (int i = 0; i < headers.Length; i++)
            {
                var c = ws.Cells[1, i + 1];
                c.Value = headers[i];
                c.Style.Font.Bold = true;
                c.Style.Font.Name = "Arial";
                c.Style.Font.Size = 11;
                c.Style.Fill.PatternType = ExcelFillStyle.Solid;
                c.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 73, 125));
                c.Style.Font.Color.SetColor(Color.White);
                c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            ws.Row(1).Height = 20;
            ws.View.FreezePanes(2, 1);
        }
    }
}
