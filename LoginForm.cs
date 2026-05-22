using System.Drawing;
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
        private readonly System.Windows.Forms.Timer _revealTimer;
        private bool _done;
        private bool _revealed;

        // How long to wait for silent re-auth (using cookies already cached in
        // WebView2's user-data folder) before actually showing the login window
        // to the user. If this elapses without finding csrf+session-id, the user
        // genuinely needs to log in, so we reveal the window.
        private const int SilentTimeoutMs = 6000;

        public string? ExtractedCookie { get; private set; }

        public LoginForm(AppConfig config)
        {
            _config = config;

            Text = "Alexa Login — Close when done";
            Width = 1100;
            Height = 750;

            // Start invisible and off-screen so a silent cookie refresh doesn't
            // flash a window on top of whatever the user is doing (e.g. a game).
            // We only show the window if SilentTimeoutMs elapses without success.
            StartPosition = FormStartPosition.Manual;
            Location = new Point(-32000, -32000);
            Opacity = 0d;
            ShowInTaskbar = false;

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

            _revealTimer = new System.Windows.Forms.Timer
            {
                Interval = SilentTimeoutMs
            };
            _revealTimer.Tick += RevealTimer_Tick;

            Load += LoginForm_Load;
        }

        private async void LoginForm_Load(object? sender, EventArgs e)
        {
            var env = await CoreWebView2Environment.CreateAsync();
            await _webView.EnsureCoreWebView2Async(env);

            var loginUrl = $"{AlexaReminderService.GetAlexaHost(_config.BaseUrl)}/api/devices-v2/device?raw=false";
            _webView.Source = new Uri(loginUrl);

            _pollTimer.Start();
            _revealTimer.Start();
        }

        private void RevealTimer_Tick(object? sender, EventArgs e)
        {
            _revealTimer.Stop();
            if (_done || _revealed)
            {
                return;
            }

            // Silent re-auth didn't work in time — the user actually needs to log in.
            _revealed = true;
            StartPosition = FormStartPosition.CenterScreen;
            var screen = Screen.FromControl(this).WorkingArea;
            Location = new Point(
                screen.Left + (screen.Width  - Width)  / 2,
                screen.Top  + (screen.Height - Height) / 2);
            ShowInTaskbar = true;
            Opacity = 1d;
            BringToFront();
            Activate();
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
                _revealTimer.Stop();

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
                _revealTimer.Dispose();
                _webView.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
