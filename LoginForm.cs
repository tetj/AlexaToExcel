using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace AlexaToExcel
{
    class LoginForm : Form
    {
        private readonly AppConfig _config;
        private readonly WebView2 _webView;
        private readonly System.Windows.Forms.Timer _pollTimer;
        private bool _done;

        public string? ExtractedCookie { get; private set; }

        public LoginForm(AppConfig config)
        {
            _config = config;

            Text = "Alexa Login — Close when done";
            Width = 1100;
            Height = 750;
            StartPosition = FormStartPosition.CenterScreen;

            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_webView);

            _pollTimer = new System.Windows.Forms.Timer
            {
                Interval = 1500
            };
            _pollTimer.Tick += PollTimer_Tick;

            Load += LoginForm_Load;
        }

        private async void LoginForm_Load(object? sender, EventArgs e)
        {
            var env = await CoreWebView2Environment.CreateAsync();
            await _webView.EnsureCoreWebView2Async(env);

            var loginUrl = $"{AlexaReminderService.GetAlexaHost(_config.BaseUrl)}/api/devices-v2/device?raw=false";
            _webView.Source = new Uri(loginUrl);

            _pollTimer.Start();
        }

        private async void PollTimer_Tick(object? sender, EventArgs e)
        {
            if (_done)
            {
                return;
            }

            try
            {
                var cookies = await _webView.CoreWebView2.CookieManager
                    .GetCookiesAsync(AlexaReminderService.GetAlexaHost(_config.BaseUrl));

                bool hasCsrf      = cookies.Any(c => c.Name.Equals("csrf",       StringComparison.OrdinalIgnoreCase));
                bool hasSessionId = cookies.Any(c => c.Name.Equals("session-id", StringComparison.OrdinalIgnoreCase));

                if (!hasCsrf || !hasSessionId)
                {
                    return;
                }

                _done = true;
                _pollTimer.Stop();

                ExtractedCookie = string.Join("; ", cookies
                    .Where(c => !string.IsNullOrWhiteSpace(c.Value))
                    .Select(c => $"{c.Name}={c.Value}"));

                DialogResult = DialogResult.OK;
                Close();
            }
            catch
            {
                // WebView2 not ready yet — will retry next tick
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _pollTimer.Dispose();
                _webView.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
