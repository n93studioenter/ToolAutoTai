using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace ToolTaiHD
{
    /// <summary>
    /// Pool các tab Selenium — 1 Chrome duy nhất, nhiều tab dùng chung
    /// </summary>
    public class ChromePool : IDisposable
    {
        private readonly IWebDriver _driver;
        private readonly object _lock = new object();
        private readonly List<string> _availableTabs = new List<string>();
        private readonly SemaphoreSlim _semaphore;
        private int _maxTabs;
        private bool _disposed;

        public ChromePool(int maxTabs = 5)
        {
            _maxTabs = maxTabs;
            _semaphore = new SemaphoreSlim(maxTabs, maxTabs);

            // Cấu hình Chrome dùng chung
            var options = new ChromeOptions();
            options.AddArgument("--disable-gpu");
            options.AddArgument("--no-sandbox");
            options.AddArgument("--disable-dev-shm-usage");
            options.AddArgument("--window-size=1920,1080");
            options.AddArgument("--disable-extensions");
            options.AddArgument("--disable-popup-blocking");
            options.AddArgument("--disable-web-security");
            options.AddArgument("--allow-running-insecure-content");
            options.AddArgument("--disable-blink-features=AutomationControlled");
            options.AddExcludedArgument("enable-automation");
            options.AddAdditionalOption("useAutomationExtension", false);
            options.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36");

            options.AddArgument("--disable-features=OptimizationHints,MediaRouter,PrivacySandboxSettings4," +
                                "PrivacySandboxAdsAPIs,InterestFeedContentSuggestions,Translate," +
                                "AcceptCHFrame,ChromeLabs,CalculateNativeWinOcclusion," +
                                "WebAuthentication,WebBluetooth,WebUsb,WebSerial,WebHID,WebNFC," +
                                "FileSystemAccessAPI,IdleDetection,SyncDisabled");
            options.AddArgument("--disable-notifications");
            options.AddArgument("--disable-infobars");
            options.AddArgument("--disable-save-password-bubble");
            options.AddArgument("--disable-translate");
            options.AddArgument("--no-first-run");
            options.AddArgument("--no-default-browser-check");
            options.AddArgument("--disable-default-apps");
            options.AddArgument("--use-fake-ui-for-media-stream");
            options.AddArgument("--use-fake-device-for-media-stream");
            options.AddArgument("--deny-permission-prompts");

            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddUserProfilePreference("download.directory_upgrade", true);
            options.AddUserProfilePreference("safebrowsing.enabled", true);
            options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
            options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
            options.AddUserProfilePreference("profile.default_content_setting_values.popups", 1);

            // Chặn permission
            options.AddUserProfilePreference("profile.default_content_setting_values.media_stream", 2);
            options.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2);
            options.AddUserProfilePreference("profile.default_content_setting_values.notifications", 2);
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);

            string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var service = ChromeDriverService.CreateDefaultService(driverDirectory);
            service.HideCommandPromptWindow = true;

            _driver = new ChromeDriver(service, options);

            // Tab đầu tiên
            _availableTabs.Add(_driver.CurrentWindowHandle);

            // Pre-open thêm tabs cho nhanh
            for (int i = 1; i < maxTabs; i++)
            {
                string newHandle = OpenNewTab();
                _availableTabs.Add(newHandle);
            }

            Console.WriteLine($"✅ ChromePool: Đã mở {_availableTabs.Count} tabs");
        }

        /// <summary>
        /// Mở tab mới trong cùng Chrome
        /// </summary>
        private string OpenNewTab()
        {
            lock (_lock)
            {
                // Mở tab mới bằng JS
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.open('about:blank', '_blank');");

                // Lấy handle mới nhất
                var handles = _driver.WindowHandles;
                return handles[handles.Count - 1];
            }
        }

        /// <summary>
        /// Mượn 1 tab từ pool — trả về handle của tab đó
        /// </summary>
        public TabSession AcquireTab()
        {
            _semaphore.Wait();

            string tabHandle;
            lock (_lock)
            {
                if (_availableTabs.Count > 0)
                {
                    tabHandle = _availableTabs[0];
                    _availableTabs.RemoveAt(0);
                }
                else
                {
                    // Mở tab mới nếu hết
                    tabHandle = OpenNewTab();
                }
            }

            // Switch sang tab đó
            lock (_lock)
            {
                _driver.SwitchTo().Window(tabHandle);
            }

            return new TabSession(this, _driver, tabHandle);
        }

        /// <summary>
        /// Trả tab về pool
        /// </summary>
        internal void ReleaseTab(string handle)
        {
            lock (_lock)
            {
                if (!_availableTabs.Contains(handle))
                {
                    // Reset tab về about:blank để dùng lại
                    try
                    {
                        _driver.SwitchTo().Window(handle);
                        _driver.Navigate().GoToUrl("about:blank");
                    }
                    catch { }

                    _availableTabs.Add(handle);
                }
            }

            _semaphore.Release();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try { _driver?.Quit(); } catch { }
            try { _driver?.Dispose(); } catch { }
            _semaphore?.Dispose();
        }
    }

    /// <summary>
    /// Session cho 1 tab — tự động trả về pool khi dispose
    /// </summary>
    public class TabSession : IDisposable
    {
        private readonly ChromePool _pool;
        private readonly string _handle;
        private bool _disposed;

        public IWebDriver Driver { get; }

        internal TabSession(ChromePool pool, IWebDriver driver, string handle)
        {
            _pool = pool;
            Driver = driver;
            _handle = handle;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _pool.ReleaseTab(_handle);
        }
    }
}