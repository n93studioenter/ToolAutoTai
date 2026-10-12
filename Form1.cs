using AutoUpdaterDotNET;
using BrotliSharpLib;
using ClosedXML.Excel;
using ClosedXML.Parser;
using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.Data.Utils;
using DevExpress.Xpo.DB.Helpers;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraPrinting.Native.Navigation;
using DevExpress.XtraPrinting.Native.WebClientUIControl;
using DevExpress.XtraWaitForm;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Emgu.CV.Util;
using FuzzySharp;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using static DevExpress.Data.Helpers.FindSearchRichParser;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using static ToolTaiHD.Form1;
using Application = System.Windows.Forms.Application;
using Color = System.Drawing.Color;
using Gray = Emgu.CV.Structure.Gray;
using Path = System.IO.Path;
using Process = System.Diagnostics.Process;
namespace ToolTaiHD
{

    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {

        private class InvoiceDownloadTask
        {
            public string PathXml { get; set; }
            public string Mst { get; set; }
            public string Shdon { get; set; }
            public string Khhd { get; set; }
            public string Ncc { get; set; }
            public string Mabimat { get; set; }
            public int Type { get; set; }
            public string RootFolder { get; set; }
            public DateTime EnqueuedAt { get; set; } = DateTime.Now;
        }

        private readonly ConcurrentQueue<InvoiceDownloadTask> _pdfQueue
            = new ConcurrentQueue<InvoiceDownloadTask>();

        private readonly SemaphoreSlim _pdfSignal = new SemaphoreSlim(0, int.MaxValue);
        private CancellationTokenSource _pdfCts;
        private Task _pdfWorkerTask;
        private volatile bool _pdfWorkerRunning;

        private const int MAX_PARALLEL_PDF = 5;   // số Chrome tối đa cùng lúc
        private const int PDF_BATCH_SIZE = 5;     // nghỉ sau mỗi 5 HĐ
        private const int PDF_BATCH_DELAY_MS = 1000;
        // ═══════════════════════════════════════════════════════════════════
        //  CONSTANTS — Tập trung hằng số, tránh magic string/number
        // ═══════════════════════════════════════════════════════════════════
        public static class AppConstants
        {
            public const string TooldbPath = @"\\192.168.1.90\Ke toan 2025 New\1 Copi vao dung 1\Hoadon\Tooldb.accdb";
            public const string VietstarAccountPath = @"\\192.168.1.90\Ke toan 2025 New\1 Copi vao dung 1\Hoadon\VietstarAccount.mdb";
            public const string DbPassword = "1@35^7*9)1";

            public static string BuildConnectionString(string dbPath) =>
                $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Jet OLEDB:Database Password={DbPassword};";

            public const string GdtBaseUrl = "https://hoadondientu.gdt.gov.vn";
            public const string GdtCaptchaUrl = GdtBaseUrl + "/api/captcha";
            public const string GdtLoginUrl = GdtBaseUrl + "/api/security-taxpayer/authenticate";
            public const string GdtProfileUrl = GdtBaseUrl + "/api/security-taxpayer/profile";
            public const string ViettelLoginUrl = "https://vinvoice.viettel.vn/api/auth/login";

            public const int DefaultHttpTimeoutSeconds = 30;
            public const int MaxRetryAttempts = 3;
            public const int MaxCaptchaRetryAttempts = 8;
            public const int ExcelCacheMinutes = 300;
            public const int MinValidExcelSize = 2048;

            public const string UserAgent =
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36";

            public const string HdVaoFolder = "HDVao";
            public const string HdRaFolder = "HDRa";
            public const string YearPrefix = "HD";

            public const string TaskBlock1 = "Block_L1";
            public const string TaskBlock2 = "Block_L2";
            public const string TaskBlock3 = "Block_L3";
            public const string StartupAppName = "Saoviet Auto";
        }

        // ═══════════════════════════════════════════════════════════════════
        //  REGEX — Compile 1 lần, dùng nhiều nơi
        // ═══════════════════════════════════════════════════════════════════
        public static class RegexPatterns
        {
            public static readonly Regex LeadingZeros = new Regex(@"^0+", RegexOptions.Compiled);
            public static readonly Regex QuantityUnit = new Regex(@"(\d+(g|ml|L|kg)|x\d+|(\d+\s*cái))",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
            public static readonly Regex SpecialChars = new Regex(@"[^\w\s]", RegexOptions.Compiled);
            public static readonly Regex MultiWhitespace = new Regex(@"\s+", RegexOptions.Compiled);
            public static readonly Regex BracketContent = new Regex(@"\(.*?\)", RegexOptions.Compiled);
            public static readonly Regex SquareBracketContent = new Regex(@"\[.*?\]", RegexOptions.Compiled);
            public static readonly Regex CurlyBraceContent = new Regex(@"\{.*?\}", RegexOptions.Compiled);
            public static readonly Regex JwtToken = new Regex(@"""(eyJ[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+)""",
                RegexOptions.Compiled);
            public static readonly Regex DiscountAmount = new Regex(@"\d{1,3}(?:\.\d{3})*(?:,\d+)?",
                RegexOptions.Compiled);
        }

        // ═══════════════════════════════════════════════════════════════════
        //  HTTP CLIENT FACTORY — Singleton tránh socket exhaustion
        // ═══════════════════════════════════════════════════════════════════
        public static class GdtHttpClientFactory
        {
            private static readonly Lazy<HttpClient> _lazyClient = new Lazy<HttpClient>(CreateClient);

            public static HttpClient Instance => _lazyClient.Value;

            private static HttpClient CreateClient()
            {
                var handler = new HttpClientHandler
                {
                    AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                    AllowAutoRedirect = true,
                    MaxConnectionsPerServer = 10
                };

                var client = new HttpClient(handler)
                {
                    Timeout = TimeSpan.FromSeconds(AppConstants.DefaultHttpTimeoutSeconds)
                };

                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", AppConstants.UserAgent);
                client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json, text/plain, */*");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Accept-Language", "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Origin", AppConstants.GdtBaseUrl);
                client.DefaultRequestHeaders.TryAddWithoutValidation("Referer", AppConstants.GdtBaseUrl + "/");
                client.DefaultRequestHeaders.TryAddWithoutValidation("sec-ch-ua", "\"Google Chrome\";v=\"140\", \"Chromium\";v=\"140\", \"Not=A?Brand\";v=\"24\"");
                client.DefaultRequestHeaders.TryAddWithoutValidation("sec-ch-ua-mobile", "?0");
                client.DefaultRequestHeaders.TryAddWithoutValidation("sec-ch-ua-platform", "\"Windows\"");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Sec-Fetch-Dest", "empty");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Sec-Fetch-Mode", "cors");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Sec-Fetch-Site", "same-origin");

                return client;
            }

            public static HttpRequestMessage CreateRequest(HttpMethod method, string url, string action = "")
            {
                var req = new HttpRequestMessage(method, url);
                req.Headers.TryAddWithoutValidation("Request-Id", Guid.NewGuid().ToString());
                req.Headers.TryAddWithoutValidation("End-Point", "/tra-cuu/tra-cuu-hoa-don");
                req.Headers.TryAddWithoutValidation("Action", string.IsNullOrEmpty(action) ? "" : Uri.EscapeDataString(action));
                return req;
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        //  MAIN FORM
        // ═══════════════════════════════════════════════════════════════════
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("kernel32.dll")]
        private static extern uint WTSGetActiveConsoleSessionId();

        #region Classes

        public class MyJson
        {
            public string Key { get; set; }
            public string Content { get; set; }
        }

        public class TokenResponse
        {
            public string token { get; set; }
        }

        public class InvoiceInfo
        {
            public string Mst { get; set; }
            public string SHHD { get; set; }
            public string Sohd { get; set; }
            public DateTime NLap { get; set; }
            public string Khhd { get; set; }
            public string DirectoryPath { get; set; }
            public string Url { get; set; }
            public int Type { get; set; } // 1: Đầu vào, 2: Đầu ra
        }

        public class Congty
        {
            public string Name { get; set; }
            public string Path { get; set; }
            public int Status { get; set; }
            public string Username { get; set; }
            public string Password { get; set; }
        }

        public class Company
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public string Dbpath { get; set; }
            public int FolderPath { get; set; }
            public string MST { get; set; }
            public int STT { get; set; }
            public string Status { get; set; }
            public int IsRun { get; set; }
            public string Dauvao { get; set; }
            public string Daura { get; set; }
            public string Saoviet { get; set; }
        }

        public class ProfileResponse
        {
            public string password_expire { get; set; }
            public int expired { get; set; }
        }

        public class UseCookie
        {
            public string __cf_bm { get; set; }
            public string JSESSIONID { get; set; }
            public string access_token { get; set; }
            public string session_token { get; set; }
        }

        public class LoginResponse
        {
            public string access_token { get; set; }
            public string token_type { get; set; }
            public string refresh_token { get; set; }
            public int expires_in { get; set; }
            public string scope { get; set; }
            public long iat { get; set; }
            public string invoice_cluster { get; set; }
            public int type { get; set; }
            public string jti { get; set; }
        }

        public class SearchData
        {
            public List<InvoiceItem> content { get; set; }
        }

        public class InvoiceItem
        {
            public long id { get; set; }
            public int invoiceType { get; set; }
            public string invoiceNumber { get; set; }
            public string invoiceSeri { get; set; }
            public DateTime createdDate { get; set; }
        }

        public class SearchResponse
        {
            public int code { get; set; }
            public string message { get; set; }
            public SearchData data { get; set; }
        }

        public class VatTuInfo
        {
            public string Ma { get; set; }
            public string DonViTinh { get; set; }
        }

        public class VatTu
        {
            public int MaSo { get; set; }
            public int MaPhanLoai { get; set; }
            public string TenMaPhanLoai { get; set; }
            public string SoHieu { get; set; }
            public string GhiChu { get; set; }
            public string TenVattu { get; set; }
            public string TenVattu2 { get; set; }
            public string DonVi { get; set; }
            public double Dongia { get; set; }
            public double SoLuong { get; set; }
            public double Dongia2 { get; set; }
            public double ThanhTien { get; set; }
            public string PTGB { get; set; }
            public double Percent { get; set; }
            public double Real { get; set; }
        }

        public class FileImport
        {
            public string Path { get; set; }
            public bool Checked { get; set; }
            public int ID { get; set; }
            public string SHDon { get; set; }
            public string KHHDon { get; set; }
            public DateTime NLap { get; set; }
            public DateTime Ngaytao { get; set; }
            public string Ten { get; set; }
            public string Noidung { get; set; }
            public string TKCo { get; set; }
            public string TKNo { get; set; }
            public int TkThue { get; set; }
            public string Mst { get; set; }
            public double TongTien { get; set; }
            public double TPhi { get; set; }
            public double TgTCThue { get; set; }
            public double TgTThue { get; set; }
            public int Vat { get; set; }
            public int Vat2 { get; set; }
            public int Vat3 { get; set; }
            public double TVat { get; set; }
            public double TVat2 { get; set; }
            public double TVat3 { get; set; }
            public int Type { get; set; }
            public bool isAcess { get; set; }
            public bool isHaschild { get; set; }
            public bool Khautruthue { get; set; }
            public int InvoiceType { get; set; }
            public string SoHieuTP { get; set; }
            public string Macdinhstatus { get; set; }
            public List<FileImportDetail> fileImportDetails;

            public class FileImportDetail
            {
                public int ID { get; set; }
                public string Ten { get; set; }
                public int ParentId { get; set; }
                public string SoHieu { get; set; }
                public double Soluong { get; set; }
                public double Dongia { get; set; }
                public string DVT { get; set; }
                public string MaCT { get; set; }
                public string TKNo { get; set; }
                public string TKCo { get; set; }
                public double TTien { get; set; }
                public double SoPSGoc { get; set; }
                public double TgTThue { get; set; }
                public string Percent { get; set; }
                public int Tchat { get; set; }
                public int IsMacdinh { get; set; }
                public int VAT { get; set; }

                public FileImportDetail(int id, string ten, int parentId, string soHieu, double soluong, double dongia,
                    string dVT, string maCT, string tkNo, string tkCo, double ttien, string percent, int tchat, int vat)
                {
                    Ten = ten; ParentId = parentId; SoHieu = soHieu; Soluong = soluong; Dongia = dongia;
                    DVT = dVT; MaCT = maCT; TKNo = tkNo; TKCo = tkCo; TTien = ttien; ID = id;
                    Percent = percent; Tchat = tchat; IsMacdinh = 0; SoPSGoc = ttien; VAT = vat;
                }
            }

            public static int Id = 1;

            public FileImport(bool chk, string path, string shdon, string khhdon, DateTime nlap, DateTime ntao,
                string ten, string noidung, string tkno, string tkco, int tkthue, string mst, double tongTien,
                int vat, int type, string tenTP, bool isacess, double tPhi, double tgTCThue, double tgTThue,
                bool _isHaschild, int _InvoiceType, double tvat, int vat2, double tvat2, int vat3, double tvat3,
                string macdinhstatus, bool khautruthue)
            {
                ID = Id; SHDon = shdon; KHHDon = khhdon; NLap = nlap; Ngaytao = ntao; Ten = ten;
                Noidung = noidung; TKCo = tkco; TKNo = tkno; TkThue = tkthue; Mst = mst; TongTien = tongTien;
                Vat = vat; Id += 1; fileImportDetails = new List<FileImportDetail>(); Type = type;
                Checked = chk; Path = path; SoHieuTP = tenTP; isAcess = isacess; TPhi = tPhi;
                TgTCThue = tgTCThue; TgTThue = tgTThue; isHaschild = _isHaschild; InvoiceType = _InvoiceType;
                TVat = tvat; Vat2 = vat2; TVat2 = tvat2; Vat3 = vat3; TVat3 = tvat3;
                Macdinhstatus = macdinhstatus; Khautruthue = khautruthue;
            }
        }

        public class Invoice
        {
            public string Pban { get; set; }
            public string Khmshdon { get; set; }
            public string Khhdon { get; set; }
            public string Shdon { get; set; }
            public string Dvtte { get; set; }
            public string Tgia { get; set; }
            public string Thtttoan { get; set; }
            public string Msttcgp { get; set; }
            public string Nbten { get; set; }
            public string Nbmst { get; set; }
            public string Nbdchi { get; set; }
            public string Nbsdthoai { get; set; }
            public string Nmten { get; set; }
            public string Nmmst { get; set; }
            public string Nmdchi { get; set; }
            public string Mkhang { get; set; }
            public List<HangHoa> Hdhhdvu { get; set; }
            public decimal? Tgtcthue { get; set; }
            public decimal? Tgtthue { get; set; }
            public decimal? Ttcktmai { get; set; }
            public decimal? Tgtttbso { get; set; }
            public string Tgtttbchu { get; set; }
            public decimal? Tgtphi { get; set; }
        }

        public class HangHoa
        {
            public string Ten { get; set; }
            public string Dvtinh { get; set; }
            public decimal? Sluong { get; set; }
            public decimal? Dgia { get; set; }
            public decimal? Tsuat { get; set; }
            public decimal? Thtien { get; set; }
        }

        public class DownloadStats
        {
            public int Downloaded { get; set; }
            public int Failed { get; set; }
            public int TimeoutCount { get; set; }
            public object LockObj { get; } = new object();
        }

        public class TTinChung
        {
            public List<VatTu> lstvt = new List<VatTu>();
            public DataTable PLHH = new DataTable();
            public DataTable tbKhachhang = new DataTable();

            public HashSet<(string Mst, string SoHD, string KyHieu, DateTime NLap, int Type)> lookupHoaDonCT { get; set; }
                = new HashSet<(string Mst, string SoHD, string KyHieu, DateTime NLap, int Type)>();

            public HashSet<(string MST, string SHDon, DateTime NLap, int Type)> lookupTbImport
                = new HashSet<(string MST, string SHDon, DateTime NLap, int Type)>();
        }

        public enum LoginResult
        {
            Success, WrongCredentials, PasswordExpired, NetworkError, CaptchaFailed
        }

        public class LoginOutcome
        {
            public LoginResult Result { get; set; }
            public string Token { get; set; } = "";
            public DateTime? PasswordExpireDate { get; set; }
            public string Message { get; set; } = "";
        }

        #endregion

        #region Fields

        public HashSet<(string Mst, string SoHD, string KyHieu, DateTime NLap, int Type)> lookupHoaDonCT { get; }
            = new HashSet<(string Mst, string SoHD, string KyHieu, DateTime NLap, int Type)>();

        private HashSet<(string MST, string SHDon, DateTime NLap, int Type)> lookupTbImport
            = new HashSet<(string MST, string SHDon, DateTime NLap, int Type)>();

        string mstcongty = "";
        int status = 0;
        string connectionString { get; set; }
        string connectionString2 { get; set; }
        string dbPath { get; set; }
        DataTable tbCompany;
        DataTable tbimport;
        DataTable tbchungtu;
        int trylogin = 0;
        public string tokken { get; set; }
        public string connectionString3 = "";

        int thangchon = 0;
        int sothutu = 1;
        bool isAddhd = true;
        string csohieu = "";
        int type = 0;
        private string MSTCongTY = "";
        private string CCCD = "";

        DataTable tbimports;
        DataTable existingTbChungtu;
        DataTable existingTbHeThongTK;
        public DataTable existingTbHoadon;
        DataTable tbNhapkhonguyenlieu;
        DataTable tbTonkho;
        DataTable tbNhapkhotp;
        DataTable tbRegister;
        DataTable tbLicense;
        DataTable ListPhanloaiVattu;
        DataTable tbDinhDanhtaikhoan = new DataTable();
        System.Data.DataTable tbDinhDanhtaikhoanUuTien;
        public DataTable tbKhachhang = new DataTable();
        public DataTable existingTbChungtu2;

        public List<VatTu> lstvt = new List<VatTu>();
        private Dictionary<string, (string TenVattu, string TenVattu2, string DonVi, double Dongia, double SoLuong)> vatTuLookup;
        private Dictionary<string, string> _lookupByTenChinh;
        private Dictionary<string, VatTuInfo> _lookupByTenChinhs;
        public Dictionary<string, string> _lookupByTenPhu;
        private Dictionary<string, string> _synonymDictionary;
        private Dictionary<string, (string TenChuan, string TenPhuChuan, string QuyCach, string DonVi, double Dongia, double soluong)> _optimizedVatTu;
        private bool _isIndexBuilt = false;
        private Dictionary<string, HashSet<string>> _keywordIndex;
        private Dictionary<string, HashSet<string>> _quyCachIndex;
        private Dictionary<string, TbImportDetail> cacheMatHangTrongHoaDon;

        private readonly Dictionary<string, (string SoHieu, double Percent)> _cacheToanCuc
            = new Dictionary<string, (string SoHieu, double Percent)>(StringComparer.OrdinalIgnoreCase);

        private BindingList<FileImport> lstImportVao = new BindingList<FileImport>();
        private BindingList<FileImport> lstImportRa = new BindingList<FileImport>();
        private List<TbImport> lstdsVao = new List<TbImport>();
        private List<TbImport> lstdsRa = new List<TbImport>();

        List<string> lstCongtyKhoa = new List<string>();

        public static LoginResponse loginResponse { get; set; } = new LoginResponse();
        public static UseCookie useCookie { get; set; } = new UseCookie();

        private NotifyIcon notifyIcon;
        private ContextMenuStrip contextMenu;
        LoadingForm loadingForm { get; set; }

        private static Action<string> _logAction;
        private bool _isExiting = false;

        #endregion

        #region Constructor & Form Load

        public bool isAutoMode { get; set; }

        public Form1(bool isAutoMode = false)
        {
            loadingForm = new LoadingForm();
            if (isAutoMode == false)
            {
                this.Opacity = 0;
                loadingForm.Show();
                Application.DoEvents();
            }
            InitializeComponent();
            timerClock.Interval = 1000;
            timerClock.Tick += TimerClock_Tick;
            timerClock.Start();

            this.isAutoMode = isAutoMode;

            if (isAutoMode)
            {
                this.ShowInTaskbar = true;
                Text = "Ứng dụng - Chạy từ Task Scheduler";
                InitializeNotifyIcon();
                status = 2;
            }
            else
            {
                Text = "Ứng dụng - Chạy thủ công";
                ShowNormalUI();
                status = 1;
            }

            _logAction = (message) =>
            {
                try
                {
                    if (richTextBox1 == null || richTextBox1.IsDisposed)
                        return;

                    if (richTextBox1.InvokeRequired)
                    {
                        richTextBox1.Invoke(new Action(() =>
                        {
                            if (!richTextBox1.IsDisposed)
                            {
                                richTextBox1.Text += $"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}";
                                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                                richTextBox1.ScrollToCaret();
                            }
                        }));
                    }
                    else
                    {
                        richTextBox1.Text += $"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                    }
                }
                catch (ObjectDisposedException) { }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi Log UI: {ex.Message}");
                }
            };
        }

        private void TimerClock_Tick(object sender, EventArgs e)
        {
            lblClock.Text = DateTime.Now.ToString("HH:mm:ss");
        }

        private void InitializeNotifyIcon()
        {
            notifyIcon = new NotifyIcon();
            try { notifyIcon.Text = "Tool tải tự động"; } catch { }

            string appPath = Application.StartupPath;
            string iconPath = Path.Combine(appPath, @"..\..\Resources\saovietfavicon.ico");

            notifyIcon.Icon = File.Exists(iconPath) ? new Icon(iconPath) : SystemIcons.Application;

            contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("Hiện ứng dụng", null, ShowApp_Click);
            contextMenu.Items.Add("-");
            contextMenu.Items.Add("Thoát", null, ExitApp_Click);
            notifyIcon.ContextMenuStrip = contextMenu;
            notifyIcon.DoubleClick += NotifyIcon_DoubleClick;
            notifyIcon.Visible = true;
        }

        private void HideToSystemTray()
        {
            Action hide = () =>
            {
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                this.Visible = false;
                notifyIcon.Visible = true;
                notifyIcon.ShowBalloonTip(3000, "Tool tải hóa đơn",
                    "Ứng dụng đang chạy ngầm trong System Tray! tháng " + comboBoxEdit1.EditValue?.ToString(),
                    ToolTipIcon.Info);
            };

            if (this.InvokeRequired) this.Invoke(hide);
            else hide();
        }

        private void NotifyIcon_DoubleClick(object sender, EventArgs e) => ShowApp();
        private void ShowApp_Click(object sender, EventArgs e) => ShowApp();

        private void ShowApp()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();
        }

        private void ExitApp_Click(object sender, EventArgs e)
        {
            notifyIcon.Visible = false;
            notifyIcon.Dispose();
            Application.Exit();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (this.WindowState == FormWindowState.Minimized && isAutoMode)
            {
                this.Hide();
                this.ShowInTaskbar = false;
                if (notifyIcon != null) notifyIcon.Visible = true;
            }
        }

        private void ShowNormalUI()
        {
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
        }

        #endregion

        #region Grid Setup

        private void SetupGridFolderColumn()
        {
            var folderButton = new RepositoryItemButtonEdit();
            folderButton.Buttons[0].Kind = ButtonPredefines.Glyph;
            folderButton.Buttons[0].ImageOptions.Image = Properties.Resources.open2_16x16;
            folderButton.Buttons[0].ToolTip = "Mở thư mục";
            gridView1.Columns["Folder"].ColumnEdit = folderButton;
            folderButton.ButtonClick += FolderButton_ButtonClick;
        }

        private void SetupGridClearColumn()
        {
            var btn = new RepositoryItemButtonEdit();
            btn.Buttons[0].Kind = ButtonPredefines.Glyph;
            btn.Buttons[0].ImageOptions.Image = Properties.Resources.cancel_32x32;
            btn.Buttons[0].ToolTip = "Clear";
            gridView1.Columns["Clear"].ColumnEdit = btn;
            btn.ButtonClick += ClearButton_ButtonClick;
        }

        private void SetupGridDeleteColumn()
        {
            var btn = new RepositoryItemButtonEdit();
            btn.Buttons[0].Kind = ButtonPredefines.Glyph;
            btn.Buttons[0].ImageOptions.Image = Properties.Resources.cancel_32x32;
            btn.Buttons[0].ToolTip = "Delete";
            gridView1.Columns["Delete"].ColumnEdit = btn;
            btn.ButtonClick += DeleteButton_ButtonClick;
        }

        private void SetupGridCheckBox()
        {
            var checkEdit = new RepositoryItemCheckEdit();
            checkEdit.ValueChecked = 1;
            checkEdit.ValueUnchecked = 0;
            gridView1.Columns["IsRun"].ColumnEdit = checkEdit;
        }

        private void DeleteButton_ButtonClick(object sender, ButtonPressedEventArgs e)
        {
            int rowHandle = gridView1.FocusedRowHandle;
            if (rowHandle < 0) return;

            try
            {
                ExecuteQueryResult("DELETE FROM [tbCompany] where MST =?",
                    new OleDbParameter("?", gridView1.GetRowCellValue(rowHandle, "MST")));

                string dbpath = gridView1.GetRowCellValue(rowHandle, "Dbpath")?.ToString();
                string conn = AppConstants.BuildConnectionString(dbpath);

                ExecuteQueryResult2("UPDATE tbRegister set IsRegisTry=? where Username=?",
                    conn,
                    new OleDbParameter("?", "0"),
                    new OleDbParameter("?", gridView1.GetRowCellValue(rowHandle, "MST")));

                string computerName = Environment.MachineName;
                tbCompany = ExecuteQuery("SELECT * FROM tbCompany WHERE Saoviet = ? order by STT",
                    new OleDbParameter("?", computerName));
                gridControl1.DataSource = tbCompany;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void ClearButton_ButtonClick(object sender, ButtonPressedEventArgs e)
        {
            int rowHandle = gridView1.FocusedRowHandle;
            if (rowHandle < 0) return;

            string dbpath = gridView1.GetRowCellValue(rowHandle, "Dbpath")?.ToString();
            string conn = AppConstants.BuildConnectionString(dbpath);
            ExecuteQueryResult2("DELETE FROM [tbimport]", conn);
            ExecuteQueryResult2("DELETE FROM [tbimportdetail]", conn);
        }

        private void FolderButton_ButtonClick(object sender, ButtonPressedEventArgs e)
        {
            int rowHandle = gridView1.FocusedRowHandle;
            if (rowHandle < 0) return;

            string folderPath = gridView1.GetRowCellValue(rowHandle, "FolderPath")?.ToString();
            if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
                Process.Start("explorer.exe", folderPath);
            else
                XtraMessageBox.Show("Thư mục không tồn tại!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void gridView1_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            try
            {
                var view = sender as GridView;
                if (view == null || e.RowHandle < 0 || e.Column == null) return;

                var rowData = view.GetRow(e.RowHandle) as DataRowView;
                if (rowData == null) return;

                int id = Convert.ToInt32(rowData["ID"] ?? 0);
                string name = rowData["Name"]?.ToString() ?? "Unknown";

                Log($"🔄 [{name}] {e.Column.FieldName} = {e.Value}");

                ExecuteQueryResult($"UPDATE tbCompany SET {e.Column.FieldName} = ? WHERE ID = ?",
                    new OleDbParameter("?", e.Value),
                    new OleDbParameter("?", id));

                string computerName = Environment.MachineName;
                tbCompany = ExecuteQuery("SELECT * FROM tbCompany where Saoviet = ? order by STT",
                    new OleDbParameter("?", computerName));
                gridControl1.DataSource = tbCompany;
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi CellValueChanged: {ex.Message}");
            }
        }

        private void gridView1_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.RowHandle < 0) return;

            object hopleVal = gridView1.GetRowCellValue(e.RowHandle, "Hople");
            if (hopleVal == null || hopleVal == DBNull.Value) return;
            if (Convert.ToInt32(hopleVal) == 1) return;
            if (e.Column.FieldName != "Dbpath") return;

            if (Convert.ToInt32(hopleVal) == -1)
            {
                e.Appearance.BackColor = Color.LightCoral;
                e.Appearance.ForeColor = Color.White;
                e.Appearance.Options.UseBackColor = true;
                //e.Appearance.UseForeColor = true;
            }
        }

        #endregion

        #region Internet Check

        private void WaitForInternetConnection()
        {
            UpdateRadioButtonText("⏳ Đang chờ kết nối Internet...", false);

            int retryCount = 0;
            while (!IsInternetAvailable())
            {
                retryCount++;
                UpdateRadioButtonText($"⏳ Đang chờ kết nối Internet... (Đã chờ {retryCount * 5} giây)", false);
                Thread.Sleep(5000);

                if (retryCount > 60)
                {
                    DialogResult result = MessageBox.Show(
                        "Không có kết nối Internet trong 5 phút!\nBạn có muốn tiếp tục chờ không?",
                        "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (result == DialogResult.No)
                    {
                        this.Invoke(new Action(() => this.Close()));
                        return;
                    }
                    retryCount = 0;
                }
            }
            UpdateRadioButtonText("✅ Đã có kết nối Internet!", true);
        }

        private void UpdateRadioButtonText(string text, bool isChecked)
        {
            if (radioButton1.InvokeRequired)
                radioButton1.Invoke(new Action(() => { radioButton1.Text = text; radioButton1.Checked = isChecked; }));
            else { radioButton1.Text = text; radioButton1.Checked = isChecked; }
        }

        public bool IsInternetAvailable()
        {
            try
            {
                using (var ping = new Ping())
                {
                    var reply = ping.Send("8.8.8.8", 3000);
                    if (reply != null && reply.Status == IPStatus.Success) return true;
                }
                using (var ping = new Ping())
                {
                    var reply = ping.Send("1.1.1.1", 3000);
                    return reply != null && reply.Status == IPStatus.Success;
                }
            }
            catch { return false; }
        }

        #endregion

        #region Database Operations

        public DataTable ExecuteQuery(string query, params OleDbParameter[] parameters)
        {
            DataTable dataTable = new DataTable();
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (parameters != null) command.Parameters.AddRange(parameters);
                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                            dataAdapter.Fill(dataTable);
                    }
                }
                catch (Exception ex) { Log($"Lỗi ExecuteQuery: {ex.Message}"); }
            }
            return dataTable;
        }

        public DataTable ExecuteQuery2(string query, string connectionst, params OleDbParameter[] parameters)
        {
            DataTable dataTable = new DataTable();
            using (OleDbConnection connection = new OleDbConnection(connectionst))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (parameters != null) command.Parameters.AddRange(parameters);
                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                            dataAdapter.Fill(dataTable);
                    }
                }
                catch (Exception ex) { Log($"Lỗi ExecuteQuery2: {ex.Message}"); }
            }
            return dataTable;
        }

        public int ExecuteQueryResult(string query, params OleDbParameter[] parameters)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    if (parameters != null) command.Parameters.AddRange(parameters);
                    command.ExecuteNonQuery();
                }
                using (OleDbCommand idCommand = new OleDbCommand("SELECT @@IDENTITY", connection))
                {
                    object result = idCommand.ExecuteScalar();
                    return Convert.ToInt32(result);
                }
            }
        }

        public int ExecuteQueryResult2(string query, string connectionst, params OleDbParameter[] parameters)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionst))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    if (parameters != null) command.Parameters.AddRange(parameters);
                    command.ExecuteNonQuery();
                }
                using (OleDbCommand idCommand = new OleDbCommand("SELECT @@IDENTITY", connection))
                {
                    object result = idCommand.ExecuteScalar();
                    return Convert.ToInt32(result);
                }
            }
        }

        public DataTable ExecuteQuery3(string query, params OleDbParameter[] parameters)
        {
            DataTable dataTable = new DataTable();
            using (OleDbConnection connection = new OleDbConnection(connectionString3))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (parameters != null) command.Parameters.AddRange(parameters);
                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                            dataAdapter.Fill(dataTable);
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
            }
            return dataTable;
        }

        #endregion

        #region Form Load

        private async void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                Application.DoEvents();
                string computerName = Environment.MachineName;
                labelControl8.Text = computerName;
                this.Text = $"{computerName} - Saoviet auto version 1.1";
                await Task.Run(() => WaitForInternetConnection());

                radioButton1.Checked = true;
                radioButton1.Text = "Có kết nối Internet!";

                dbPath = AppConstants.TooldbPath;
                connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";

                tbCompany = ExecuteQuery("SELECT * FROM tbCompany WHERE Saoviet = ? order by STT",
                    new OleDbParameter("?", computerName));
                tbCompany.Columns.Add("VbCoche", typeof(string));
                tbCompany.Columns.Add("VbCoche2", typeof(string));
                tbCompany.Columns.Add("Hople", typeof(int));

                Xulydata();

                foreach (DataRow row in tbCompany.Rows)
                {
                    string vbdbpath = row["Dbpath"].ToString();
                    string connectionString2 = AppConstants.BuildConnectionString(vbdbpath);
                    string qrkh = "SELECT * FROM tbRegister";

                    try
                    {
                        var tbRegister = ExecuteQuery2(qrkh, connectionString2);
                        if (tbRegister != null && tbRegister.Rows.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(tbRegister.Rows[0]["VbCoche"].ToString()))
                            {
                                if (tbRegister.Rows[0]["VbCoche"].ToString() == "1") row["VbCoche"] = "Tự động toàn bộ";
                                if (tbRegister.Rows[0]["VbCoche"].ToString() == "2") row["VbCoche"] = "Theo mật định";
                            }
                            if (!string.IsNullOrEmpty(tbRegister.Rows[0]["VbCoche2"].ToString()))
                            {
                                if (tbRegister.Rows[0]["VbCoche2"].ToString() == "1") row["VbCoche2"] = "Tự động toàn bộ";
                                if (tbRegister.Rows[0]["VbCoche2"].ToString() == "2") row["VbCoche2"] = "Theo mật định";
                            }
                            row["Password"] = tbRegister.Rows[0]["Password"].ToString();
                            try
                            {
                                if (row["DateExpert"] != null && string.IsNullOrEmpty(row["DateExpert"].ToString()))
                                    row["DateExpert"] = tbRegister.Rows[0]["DateExpert"].ToString();
                            }
                            catch { }
                        }
                    }
                    catch { continue; }
                }

                gridControl1.DataSource = tbCompany;
                SetupGridCheckBox();
                SetupGridFolderColumn();
                SetupGridClearColumn();
                SetupGridDeleteColumn();

                string qrst = @"SELECT * FROM tbsetting WHERE Saoviet = ?";
                DataTable tbSetting = ExecuteQuery(qrst, new OleDbParameter("?", computerName));
                if (tbSetting.Rows.Count > 0)
                {
                    int autotai = tbSetting.Rows[0].Field<int?>("AutoTai") ?? 0;
                    string Block1 = tbSetting.Rows[0].Field<string>("Block1");
                    string Block2 = tbSetting.Rows[0].Field<string>("Block2");
                    string Block3 = tbSetting.Rows[0].Field<string>("Block3");

                    checkEdit1.Checked = autotai == 1;
                    chkMoc1.Checked = !string.IsNullOrEmpty(Block1);
                    chkMoc2.Checked = !string.IsNullOrEmpty(Block2);
                    chkMoc3.Checked = !string.IsNullOrEmpty(Block3);
                    txtSoluongtai.Text = tbSetting.Rows[0].Field<int>("Soluongtai").ToString();
                    txttimeout.Text = tbSetting.Rows[0].Field<int>("Timeout").ToString();

                    if (chkMoc1.Checked) { txtBlock1.Enabled = true; txtBlock1.Text = Block1; }
                    if (chkMoc2.Checked) { txtBlock2.Enabled = true; txtBlock2.Text = Block2; }
                    if (chkMoc3.Checked) { txtBlock3.Enabled = true; txtBlock3.Text = Block3; }
                }

                comboBoxEdit1.Properties.Items.Clear();
                for (int i = 1; i <= 12; i++) comboBoxEdit1.Properties.Items.Add(i);

                try
                {
                    string checkQuery = "SELECT COUNT(*) FROM tbRemember WHERE Saoviet = ?";
                    DataTable dt = ExecuteQuery(checkQuery, new OleDbParameter("?", computerName));
                    int count = Convert.ToInt32(dt.Rows[0][0]);

                    if (count == 0)
                    {
                        comboBoxEdit1.SelectedIndex = DateTime.Now.Month - 1;
                        Kichhoatngay();
                    }
                    else
                    {
                        if (status == 2)
                        {
                            DataTable dts = ExecuteQuery("SELECT * FROM tbRemember WHERE Saoviet = ?",
                                new OleDbParameter("?", computerName));
                            thangchon = int.Parse(dts.Rows[0]["ThangCT"].ToString());
                            comboBoxEdit1.SelectedIndex = thangchon - 1;
                            dateEdit1.DateTime = DateTime.Parse(dts.Rows[0]["TuNgay"].ToString());
                            dateEdit2.DateTime = DateTime.Parse(dts.Rows[0]["DenNgay"].ToString());
                        }
                        else
                        {
                            thangchon = DateTime.Now.Month;
                            comboBoxEdit1.SelectedIndex = DateTime.Now.Month - 1;
                            Kichhoatngay();
                        }
                    }
                }
                catch (Exception ex) { XtraMessageBox.Show(ex.Message); }

                if (isAutoMode)
                {
                    btnRun.PerformClick();
                    HideToSystemTray();
                }
            }
            catch (Exception ex) { XtraMessageBox.Show(ex.Message); }

            if (isAutoMode == false)
            {
                this.Opacity = 1;
                loadingForm.Close();
            }
            StartPdfWorker();
        }

        private void Kichhoatngay()
        {
            int month = int.Parse(comboBoxEdit1.EditValue != "" ? comboBoxEdit1.EditValue.ToString() : DateTime.Now.Month.ToString());
            DateTime firstDayOfMonth = new DateTime(DateTime.Now.Year, month, 1);
            DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
            if (lastDayOfMonth.Month == DateTime.Now.Month) lastDayOfMonth = DateTime.Now.Date;
            dateEdit1.DateTime = firstDayOfMonth;
            dateEdit2.DateTime = lastDayOfMonth;
        }

        #endregion

        #region Xulydata — Làm sạch dữ liệu công ty

        private void Xulydata()
        {
            if (!tbCompany.Columns.Contains("Hople"))
                tbCompany.Columns.Add("Hople", typeof(int));

            foreach (DataRow row in tbCompany.Rows)
            {
                row["Hople"] = 1;
                string dbpath = row["Dbpath"]?.ToString() ?? "";

                if (!File.Exists(dbpath))
                {
                    row["Hople"] = -1;
                    row["IsRun"] = 0;
                    ExecuteQueryResult("UPDATE tbCompany SET IsRun = ? WHERE ID = ?",
                        new OleDbParameter("?", "0"),
                        new OleDbParameter("?", row["ID"].ToString()));
                    continue;
                }

                string folderpath = row["FolderPath"]?.ToString() ?? "";
                string filePath = Path.Combine(folderpath, "vbPath.txt");
                if (!File.Exists(filePath)) continue;

                try
                {
                    string content = File.ReadAllText(filePath);
                    if (!File.Exists(content) || content == dbpath) continue;

                    string conn3 = AppConstants.BuildConnectionString(content);
                    var kq = ExecuteQuery2("SELECT * FROM tbRegister", conn3);
                    if (kq.Rows.Count == 0) continue;

                    string newMst = kq.Rows[0]["Username"].ToString();
                    string newName = Path.GetFileNameWithoutExtension(content);

                    row["Hople"] = 2;
                    row["DbPath"] = content;
                    row["MST"] = newMst;
                    row["Name"] = newName;

                    ExecuteQueryResult("UPDATE tbCompany SET DbPath=?, MST=?, Name=? WHERE ID=?",
                        new OleDbParameter("?", content),
                        new OleDbParameter("?", newMst),
                        new OleDbParameter("?", newName),
                        new OleDbParameter("?", row["ID"].ToString()));
                }
                catch { continue; }
            }
            tbCompany.AcceptChanges();
        }

        #endregion

        #region Token & Authentication — Async, không deadlock

        /// <summary>
        /// Lấy token GDT — async hoàn toàn, không block UI thread
        /// </summary>
        public async Task<LoginOutcome> GettokkenAsync(
            string username, string password, DataRow rowCopy = null, CancellationToken ct = default)
        {
            var outcome = new LoginOutcome();

            for (int retry = 1; retry <= AppConstants.MaxRetryAttempts; retry++)
            {
                ct.ThrowIfCancellationRequested();

                try
                {
                    var client = GdtHttpClientFactory.Instance;

                    // 1. Lấy captcha
                    string captchaBody = await FetchCaptchaAsync(client, ct);
                    if (string.IsNullOrEmpty(captchaBody))
                    {
                        outcome.Result = LoginResult.CaptchaFailed;
                        outcome.Message = "Không lấy được captcha";
                        continue;
                    }

                    var myJson = JsonConvert.DeserializeObject<MyJson>(captchaBody);
                    if (myJson == null || string.IsNullOrEmpty(myJson.Key) || string.IsNullOrEmpty(myJson.Content))
                    {
                        outcome.Result = LoginResult.CaptchaFailed;
                        outcome.Message = "Captcha response thiếu key/content";
                        continue;
                    }

                    // 2. Giải captcha
                    string svgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.svg");
                    await Task.Run(() => File.WriteAllText(svgPath, myJson.Content), ct);

                    string captchaResult = (await Task.Run(() =>
                        new SvgCaptchaSolver().SolveCaptcha(svgPath), ct))?.Trim() ?? "";

                    if (captchaResult.Length < 4)
                    {
                        Log($"Captcha giải sai: '{captchaResult}'");
                        await Task.Delay(500, ct);
                        continue;
                    }

                    // 3. Login
                    var loginResult = await LoginGdtAsync(client, username, password, captchaResult, myJson.Key, ct);

                    if (loginResult.Result == LoginResult.WrongCredentials)
                    {
                        lstCongtyKhoa.Add(username);
                        SafeShowMessage("Tên đăng nhập hoặc mật khẩu không đúng");
                        outcome.Result = LoginResult.WrongCredentials;
                        return outcome;
                    }

                    if (loginResult.Result == LoginResult.NetworkError)
                    {
                        Log($"⚠️ Login lỗi lần {retry}: {loginResult.Message}");
                        await Task.Delay(1000 * retry, ct);
                        continue;
                    }

                    if (!string.IsNullOrEmpty(loginResult.Token))
                    {
                        this.tokken = loginResult.Token;
                        outcome.Token = loginResult.Token;

                        // 4. Kiểm tra profile (hạn mật khẩu)
                        var expireDate = await CheckProfileAsync(client, loginResult.Token, ct);
                        if (expireDate.HasValue)
                        {
                            outcome.PasswordExpireDate = expireDate;
                            if (expireDate.Value < DateTime.Now)
                            {
                                lstCongtyKhoa.Add(username);
                                outcome.Result = LoginResult.PasswordExpired;
                                SafeShowMessage($"Mật khẩu đã hết hạn ngày {expireDate:dd/MM/yyyy}.");
                                return outcome;
                            }
                            UpdateDateExpert(rowCopy, expireDate.Value);
                            WarnPasswordExpiry(expireDate.Value);
                        }

                        outcome.Result = LoginResult.Success;
                        Log($"✅ Token: {outcome.Token.Substring(0, Math.Min(20, outcome.Token.Length))}...");
                        return outcome;
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    Log($"⚠️ Lỗi lấy token lần {retry}: {ex.Message}");
                    if (retry >= AppConstants.MaxRetryAttempts) break;
                    await Task.Delay(1000 * retry, ct);
                }
            }

            outcome.Result = LoginResult.NetworkError;
            outcome.Message = "Đăng nhập thất bại sau nhiều lần thử";
            return outcome;
        }

        private async Task<string> FetchCaptchaAsync(HttpClient client, CancellationToken ct)
        {
            for (int i = 0; i < AppConstants.MaxCaptchaRetryAttempts; i++)
            {
                try
                {
                     var req = GdtHttpClientFactory.CreateRequest(HttpMethod.Get, AppConstants.GdtCaptchaUrl);
                     var resp = await client.SendAsync(req, ct);
                    if (!resp.IsSuccessStatusCode) continue;

                    string body = await resp.Content.ReadAsStringAsync();
                    if (body.TrimStart().StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase)) continue;
                    if (body.Length > 20) return body;
                }
                catch (Exception ex) { Log($"Captcha lỗi lần {i + 1}: {ex.Message}"); }
                await Task.Delay(500, ct);
            }
            return null;
        }

        private async Task<LoginOutcome> LoginGdtAsync(
            HttpClient client, string username, string password,
            string cValue, string cKey, CancellationToken ct)
        {
            var outcome = new LoginOutcome();
            try
            {
                var payload = new { username, password, cvalue = cValue, ckey = cKey };
                string json = JsonConvert.SerializeObject(payload);

                 var req = GdtHttpClientFactory.CreateRequest(HttpMethod.Post, AppConstants.GdtLoginUrl);
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");

                 var resp = await client.SendAsync(req, ct);
                string body = await resp.Content.ReadAsStringAsync();

                if (body.TrimStart().StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase))
                {
                    outcome.Result = LoginResult.NetworkError;
                    outcome.Message = "WAF trả HTML";
                    return outcome;
                }

                if (resp.StatusCode == HttpStatusCode.Unauthorized)
                {
                    outcome.Result = LoginResult.NetworkError;
                    outcome.Message = "401 Unauthorized";
                    return outcome;
                }

                try
                {
                    var jb = JObject.Parse(body);
                    if (jb["message"]?.ToString() == "Tên đăng nhập hoặc mật khẩu không đúng")
                    {
                        outcome.Result = LoginResult.WrongCredentials;
                        return outcome;
                    }
                }
                catch { }

                if (!resp.IsSuccessStatusCode)
                {
                    outcome.Result = LoginResult.NetworkError;
                    outcome.Message = $"HTTP {(int)resp.StatusCode}";
                    return outcome;
                }

                var tokenResp = JsonConvert.DeserializeObject<TokenResponse>(body);
                if (!string.IsNullOrEmpty(tokenResp?.token))
                {
                    outcome.Token = tokenResp.token;
                    outcome.Result = LoginResult.Success;
                }
                else
                {
                    var m = RegexPatterns.JwtToken.Match(body);
                    if (m.Success)
                    {
                        outcome.Token = m.Groups[1].Value;
                        outcome.Result = LoginResult.Success;
                    }
                }
            }
            catch (Exception ex)
            {
                outcome.Result = LoginResult.NetworkError;
                outcome.Message = ex.Message;
            }
            return outcome;
        }

        private async Task<DateTime?> CheckProfileAsync(HttpClient client, string token, CancellationToken ct)
        {
            try
            {
                 var req = GdtHttpClientFactory.CreateRequest(HttpMethod.Get, AppConstants.GdtProfileUrl);
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                 var resp = await client.SendAsync(req, ct);
                if (!resp.IsSuccessStatusCode) return null;

                string body = await resp.Content.ReadAsStringAsync();
                var prof = JsonConvert.DeserializeObject<ProfileResponse>(body);

                if (!string.IsNullOrEmpty(prof?.password_expire)
                    && DateTime.TryParse(prof.password_expire, out var d))
                    return d;
            }
            catch (Exception ex) { Log("Không kiểm tra được hạn mật khẩu: " + ex.Message); }
            return null;
        }

        private static void WarnPasswordExpiry(DateTime expireDate)
        {
            var remain = expireDate - DateTime.Now;
            if (remain.TotalDays <= 3)
                SafeShowMessageStatic($"⚠ Mật khẩu sẽ hết hạn sau {remain.Days} ngày!\nNgày: {expireDate:dd/MM/yyyy}", "Cảnh báo");
            else if (remain.TotalDays <= 7)
                SafeShowMessageStatic($"Mật khẩu sắp hết hạn {expireDate:dd/MM/yyyy} (còn {remain.Days} ngày)");
        }

        private void SafeShowMessage(string message, string caption = "Thông báo")
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) BeginInvoke(new Action(() => XtraMessageBox.Show(message, caption)));
            else XtraMessageBox.Show(message, caption);
        }

        private static void SafeShowMessageStatic(string message, string caption = "Thông báo")
        {
            var form = Application.OpenForms.OfType<Form1>().FirstOrDefault();
            if (form != null) form.SafeShowMessage(message, caption);
        }

        #endregion

        #region CompanyHttpClient

        public class CompanyHttpClient : IDisposable
        {
            private readonly HttpClient _httpClient;
            private readonly string _companyName;
            private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);
            private bool _disposed;

            public CompanyHttpClient(string token, string companyName)
            {
                _companyName = companyName;

                var handler = new HttpClientHandler
                {
                    AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                    AllowAutoRedirect = true,
                    MaxConnectionsPerServer = 4
                };

                _httpClient = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(120) };

                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/octet-stream"));
                _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", AppConstants.UserAgent);
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Accept-Language", "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7");
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Origin", AppConstants.GdtBaseUrl);
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Referer", AppConstants.GdtBaseUrl + "/tra-cuu/tra-cuu-hoa-don");
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("sec-ch-ua", "\"Google Chrome\";v=\"140\", \"Chromium\";v=\"140\", \"Not=A?Brand\";v=\"24\"");
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("sec-ch-ua-mobile", "?0");
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("sec-ch-ua-platform", "\"Windows\"");
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Sec-Fetch-Dest", "empty");
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Sec-Fetch-Mode", "cors");
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Sec-Fetch-Site", "same-origin");
            }

            private void AddGdtHeaders(HttpRequestMessage request, string endPoint = "/tra-cuu/tra-cuu-hoa-don", string action = "")
            {
                request.Headers.TryAddWithoutValidation("Request-Id", Guid.NewGuid().ToString());
                request.Headers.TryAddWithoutValidation("End-Point", endPoint);
                request.Headers.TryAddWithoutValidation("Action", string.IsNullOrEmpty(action) ? "" : Uri.EscapeDataString(action));
            }

            public async Task<HttpResponseMessage> GetAsync(string url, string action = "")
            {
                await _semaphore.WaitAsync();
                try
                {
                    using (var request = new HttpRequestMessage(HttpMethod.Get, url))
                    {
                        AddGdtHeaders(request, action: action);
                        return await _httpClient.SendAsync(request);
                    }
                }
                finally { _semaphore.Release(); }
            }

            public async Task<bool> DownloadFileAsync(string url, string filePath, int timeoutSeconds,
                CancellationToken ct = default)
            {
                for (int retry = 1; retry <= AppConstants.MaxRetryAttempts; retry++)
                {
                    ct.ThrowIfCancellationRequested();

                    try
                    {
                        using (var cts = CancellationTokenSource.CreateLinkedTokenSource(ct))
                        {
                            cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));

                            using (var request = new HttpRequestMessage(HttpMethod.Get, url))
                            {
                                AddGdtHeaders(request);

                                var response = await _httpClient.SendAsync(request,
                                    HttpCompletionOption.ResponseHeadersRead, cts.Token);

                                if ((int)response.StatusCode == 429)
                                {
                                    Log($"⚠️ {_companyName}: Rate limit! Chờ 5s...");
                                    await Task.Delay(5000, ct);
                                    continue;
                                }

                                if (response.StatusCode == HttpStatusCode.Unauthorized)
                                {
                                    Log($"❌ {_companyName}: Token hết hạn!");
                                    return false;
                                }

                                if (response.StatusCode == HttpStatusCode.Forbidden)
                                {
                                    string errBody = await response.Content.ReadAsStringAsync();
                                    bool isHtml = errBody.TrimStart().StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase);
                                    Log($"❌ {_companyName}: Bị 403 {(isHtml ? "(WAF/HTML)" : "")} lần {retry}");
                                    if (retry < AppConstants.MaxRetryAttempts)
                                    {
                                        await Task.Delay(retry * 2500, ct);
                                        continue;
                                    }
                                    return false;
                                }

                                if (response.IsSuccessStatusCode)
                                {
                                    string dir = Path.GetDirectoryName(filePath);
                                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

                                    using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
                                        await response.Content.CopyToAsync(fs);

                                    var fi = new FileInfo(filePath);
                                    if (fi.Length < 1024)
                                    {
                                        Log($"⚠️ {_companyName}: File quá nhỏ ({fi.Length} bytes)");
                                        try { File.Delete(filePath); } catch { }
                                        if (retry < AppConstants.MaxRetryAttempts)
                                        {
                                            await Task.Delay(retry * 2000, ct);
                                            continue;
                                        }
                                        return false;
                                    }
                                    return true;
                                }

                                if (response.StatusCode == HttpStatusCode.InternalServerError)
                                    Log($"⚠️ {_companyName}: Lỗi 500 lần {retry}");
                                else
                                    Log($"⚠️ {_companyName}: HTTP {(int)response.StatusCode} lần {retry}");
                            }
                        }
                    }
                    catch (OperationCanceledException) when (!ct.IsCancellationRequested)
                    {
                        Log($"⏰ {_companyName}: Timeout lần {retry}/{AppConstants.MaxRetryAttempts}");
                    }
                    catch (Exception ex)
                    {
                        Log($"⚠️ {_companyName}: Lỗi lần {retry}: {ex.Message}");
                    }

                    if (retry < AppConstants.MaxRetryAttempts)
                        await Task.Delay(retry * 2000, ct);
                }
                return false;
            }

            public async Task<string> GetStringAsync(string url, string action = "")
            {
                var response = await GetAsync(url, action);
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsStringAsync();
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                _httpClient?.Dispose();
                _semaphore?.Dispose();
            }
        }

        #endregion

        #region Main Processing

        private void UpdateStatusOnUI(DataRow row, string statusText)
        {
            try
            {
                if (row == null) return;

                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() =>
                    {
                        try
                        {
                            if (!row.Table.Columns.Contains("Status"))
                                row.Table.Columns.Add("Status", typeof(string));
                            row["Status"] = statusText;
                            gridControl1.RefreshDataSource();
                        }
                        catch (Exception ex) { Log($"⚠️ Lỗi update status: {ex.Message}"); }
                    }));
                }
                else
                {
                    if (!row.Table.Columns.Contains("Status"))
                        row.Table.Columns.Add("Status", typeof(string));
                    row["Status"] = statusText;
                    gridControl1.RefreshDataSource();
                }
            }
            catch (Exception ex) { Log($"⚠️ Lỗi UpdateStatusOnUI: {ex.Message}"); }
        }

        private void UpdateRunCountOnUI(DataRow row, int currentRun, int totalRuns)
        {
            try
            {
                if (row == null) return;

                Action action = () =>
                {
                    try
                    {
                        if (!row.Table.Columns.Contains("RunCount"))
                            row.Table.Columns.Add("RunCount", typeof(string));
                        row["RunCount"] = $"{currentRun}/{totalRuns}";
                        gridControl1.RefreshDataSource();
                    }
                    catch (Exception ex) { Log($"⚠️ Lỗi update RunCount: {ex.Message}"); }
                };

                if (this.InvokeRequired) this.Invoke(action);
                else action();
            }
            catch (Exception ex) { Log($"⚠️ Lỗi UpdateRunCountOnUI: {ex.Message}"); }
        }

        private void UpdateDateExpert(DataRow row, DateTime date)
        {
            try
            {
                if (row == null) return;

                Action action = () =>
                {
                    try
                    {
                        if (!row.Table.Columns.Contains("DateExpert"))
                            row.Table.Columns.Add("DateExpert", typeof(string));
                        row["DateExpert"] = date;

                        ExecuteQueryResult("UPDATE tbCompany SET DateExpert = ? WHERE ID = ?",
                            new OleDbParameter("?", date.ToString()),
                            new OleDbParameter("?", row["ID"].ToString()));
                        gridControl1.RefreshDataSource();
                    }
                    catch (Exception ex) { Log($"⚠️ Lỗi update DateExpert: {ex.Message}"); }
                };

                if (this.InvokeRequired) this.Invoke(action);
                else action();
            }
            catch (Exception ex) { Log($"⚠️ Lỗi UpdateDateExpert: {ex.Message}"); }
        }

        private async Task TaihoadonCongty(string vbdbpath, DataRow dtrow)
        {
            if (string.IsNullOrEmpty(vbdbpath))
            {
                Log($"⚠️ Dbpath trống cho {dtrow["Name"]}");
                return;
            }

            string conn2 = AppConstants.BuildConnectionString(vbdbpath);

            try
            {
                var getResgistry = ExecuteQuery2("SELECT * FROM tbRegister", conn2);
                if (getResgistry.Rows.Count == 0)
                {
                    Log($"⚠️ Không tìm thấy tbRegister cho {dtrow["Name"]}");
                    return;
                }

                string companyName = dtrow["Name"].ToString();
                string username = getResgistry.Rows[0]["Username"].ToString();
                string password = "";

                try
                {
                    connectionString3 = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AppConstants.VietstarAccountPath};";
                    var tbImportdetails = ExecuteQuery3("SELECT * FROM tbcompany where Username = ?",
                        new OleDbParameter("?", username));
                    if (tbImportdetails.Rows.Count > 0)
                        password = tbImportdetails.Rows[0]["Password"].ToString();
                    else
                        password = getResgistry.Rows[0]["Password"].ToString();
                }
                catch { password = getResgistry.Rows[0]["Password"].ToString(); }

                string savedPath = getResgistry.Rows[0]["Hoadonpath"].ToString();
                mstcongty = username;

                Log($"[{companyName}] ▶️ Bắt đầu xử lý...");

                var loginOutcome = await GetTokenForCompanyAsync(username, password, dtrow);
                if (loginOutcome.Result != LoginResult.Success || string.IsNullOrEmpty(loginOutcome.Token))
                {
                    Log($"❌ {companyName}: Đăng nhập thất bại!");
                    return;
                }

                using (var companyClient = new CompanyHttpClient(loginOutcome.Token, companyName))
                {
                    // BƯỚC 1: Excel đầu vào
                    Log($"[{companyName}] 📥 Bắt đầu tải Excel đầu vào...");
                    bool t1 = await XulyexelvaoAsync(companyClient, 1, savedPath, companyName, username);
                    bool t2 = await XulyexelvaoAsync(companyClient, 2, savedPath, companyName, username);
                    bool t3 = await XulyexelvaoAsync(companyClient, 3, savedPath, companyName, username);

                    if (t1 && t2 && t3) Log($"✅ {companyName}: Tải Excel đầu vào thành công!");
                    else Log($"⚠️ {companyName}: Có lỗi khi tải Excel đầu vào!");

                    // BƯỚC 2: Hóa đơn đầu vào
                    Log($"[{companyName}] 📥 Bắt đầu tải hóa đơn đầu vào...");
                    var invoicesVao = await GetListHoaDonCanTai(username, savedPath, 1);
                    if (invoicesVao.Count > 0)
                        await TaiHangLoatHoaDon(companyClient, invoicesVao, "đầu vào", companyName, dtrow);
                    else
                        Log($"📭 {companyName}: Không có hóa đơn đầu vào mới");

                    // BƯỚC 3: Excel đầu ra
                    Log($"[{companyName}] 📤 Bắt đầu tải Excel đầu ra...");
                    bool t1r = await XulyexelraAsync(companyClient, 1, savedPath, companyName, username);
                    bool t2r = await XulyexelraAsync(companyClient, 2, savedPath, companyName, username);

                    if (t1r && t2r) Log($"✅ {companyName}: Tải Excel đầu ra thành công!");
                    else Log($"⚠️ {companyName}: Có lỗi khi tải Excel đầu ra!");

                    // BƯỚC 4: Hóa đơn đầu ra
                    Log($"[{companyName}] 📤 Bắt đầu tải hóa đơn đầu ra...");
                    var tbInvoiceInfo = ExecuteQuery2("SELECT * FROM tbInvoiceInfo", conn2);
                    var tbRrtry = ExecuteQuery2("SELECT * FROM tbRegister", conn2);
                    var invoicesRa = await GetListHoaDonCanTai(username, savedPath, 2);

                    if (invoicesRa.Count > 0 || tbInvoiceInfo.Rows.Count > 0)
                    {
                        if (tbInvoiceInfo.Rows.Count > 0 && tbRrtry.Columns.Contains("IsNCC") && tbRrtry.Rows[0]["IsNCC"].ToString() == "1")
                        {
                            await Tainhacungcap(vbdbpath, dtrow);
                        }
                        else
                        {
                            await TaiHangLoatHoaDon(companyClient, invoicesRa, "đầu ra", companyName, dtrow);
                        }
                    }
                    else
                    {
                        Log($"📭 {companyName}: Không có hóa đơn đầu ra mới");
                    }
                }

                Log($"✅ {companyName}: Hoàn thành xử lý!");
            }
            catch (Exception ex)
            {
               // Log($"❌ {companyName}: Lỗi TaihoadonCongty: {ex.Message}");
            }
        }

        private async Task<LoginOutcome> GetTokenForCompanyAsync(string username, string password, DataRow rowCopy)
        {
            try
            {
                return await GettokkenAsync(username, password, rowCopy);
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi lấy token: {ex.Message}");
                return new LoginOutcome { Result = LoginResult.NetworkError, Message = ex.Message };
            }
        }

        #endregion

        #region Excel Download

        public async Task<bool> XulyexelvaoAsync(CompanyHttpClient client, int _type, string savedPath,
            string companyName, string mstcongty)
        {
            try
            {
                int month = int.Parse(comboBoxEdit1.EditValue.ToString());
                DateTime dtFrom = new DateTime(DateTime.Now.Year, month, 1);
                DateTime dtTo = dtFrom.AddMonths(1).AddDays(-1);

                string formattedDate1 = dtFrom.ToString("dd/MM/yyyy") + "T00:00:00";
                string formattedDate2 = dtTo.ToString("dd/MM/yyyy") + "T23:59:59";

                string url, filename, fileType;

                switch (_type)
                {
                    case 1:
                        url = $"{AppConstants.GdtBaseUrl}/api/query/invoices/export-excel-sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2};ttxly==5&type=purchase";
                        filename = $"{mstcongty}_HDDienTuDaCapMa.xlsx";
                        fileType = "hóa đơn điện tử có mã";
                        break;
                    case 2:
                        url = $"{AppConstants.GdtBaseUrl}/api/query/invoices/export-excel-sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2};ttxly==6&type=purchase";
                        filename = $"{mstcongty}_HDDienTuKhongMa.xlsx";
                        fileType = "hóa đơn điện tử không mã";
                        break;
                    case 3:
                        url = $"{AppConstants.GdtBaseUrl}/api/sco-query/invoices/export-excel-sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2};ttxly==8&type=purchase";
                        filename = $"{mstcongty}_HDDienTuMayTinhTien.xlsx";
                        fileType = "hóa đơn máy tính tiền";
                        break;
                    default: return false;
                }

                string currentYear = $"{AppConstants.YearPrefix}{DateTime.Now.Year}";
                string directoryPath = Path.Combine(savedPath, currentYear, AppConstants.HdVaoFolder, month.ToString());
                Directory.CreateDirectory(directoryPath);
                string filePath = Path.Combine(directoryPath, filename);

                if (File.Exists(filePath))
                {
                    TimeSpan ts = DateTime.Now - File.GetLastWriteTime(filePath);
                    if (ts.TotalMinutes < AppConstants.ExcelCacheMinutes)
                    {
                        Log($"✅ {companyName}: Đã có file {fileType} (còn mới)");
                        return true;
                    }
                    File.Delete(filePath);
                }

                Log($"📥 {companyName}: Đang tải {fileType} đầu vào...");
                bool success = await client.DownloadFileAsync(url, filePath, int.Parse(txttimeout.Text));

                if (success)
                {
                    var fi = new FileInfo(filePath);
                    if (fi.Length < AppConstants.MinValidExcelSize)
                    {
                        Log($"❌ {companyName}: File {fileType} quá nhỏ ({fi.Length} bytes)");
                        try { File.Delete(filePath); } catch { }
                        return false;
                    }
                    Log($"✅ {companyName}: Đã tải xong {fileType} đầu vào");
                    return true;
                }
                Log($"❌ {companyName}: Tải {fileType} đầu vào thất bại!");
                return false;
            }
            catch (Exception ex)
            {
                Log($"❌ {companyName}: Lỗi XulyexelvaoAsync: {ex.Message}");
                return false;
            }
        }

        public async Task<bool> XulyexelraAsync(CompanyHttpClient client, int _type, string savedPath,
            string companyName, string mstcongtys)
        {
            try
            {
                int month = int.Parse(comboBoxEdit1.EditValue.ToString());
                DateTime dtFrom = new DateTime(DateTime.Now.Year, month, 1);
                DateTime dtTo = dtFrom.AddMonths(1).AddDays(-1);

                string formattedDate1 = dtFrom.ToString("dd/MM/yyyy") + "T00:00:00";
                string formattedDate2 = dtTo.ToString("dd/MM/yyyy") + "T23:59:59";

                string url, filename, fileType;

                switch (_type)
                {
                    case 1:
                        url = $"{AppConstants.GdtBaseUrl}/api/query/invoices/export-excel?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2}";
                        filename = $"{mstcongtys}_Hoadondientu.xlsx";
                        fileType = "hóa đơn điện tử (đầu ra)";
                        break;
                    case 2:
                        url = $"{AppConstants.GdtBaseUrl}/api/sco-query/invoices/export-excel?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2}";
                        filename = $"{mstcongtys}_HDDienTuMayTinhTien.xlsx";
                        fileType = "hóa đơn máy tính tiền (đầu ra)";
                        break;
                    default: return false;
                }

                string currentYear = $"{AppConstants.YearPrefix}{DateTime.Now.Year}";
                string directoryPath = Path.Combine(savedPath, currentYear, AppConstants.HdRaFolder, month.ToString());
                Directory.CreateDirectory(directoryPath);
                string filePath = Path.Combine(directoryPath, filename);

                if (File.Exists(filePath))
                {
                    TimeSpan ts = DateTime.Now - File.GetLastWriteTime(filePath);
                    if (ts.TotalMinutes < AppConstants.ExcelCacheMinutes)
                    {
                        Log($"✅ {companyName}: Đã có file {fileType} (còn mới)");
                        return true;
                    }
                    File.Delete(filePath);
                }

                Log($"📤 {companyName}: Đang tải {fileType}...");
                bool success = await client.DownloadFileAsync(url, filePath, int.Parse(txttimeout.Text));

                if (success)
                {
                    var fi = new FileInfo(filePath);
                    if (fi.Length < AppConstants.MinValidExcelSize)
                    {
                        Log($"❌ {companyName}: File {fileType} quá nhỏ ({fi.Length} bytes)");
                        try { File.Delete(filePath); } catch { }
                        return false;
                    }
                    Log($"✅ {companyName}: Đã tải xong {fileType}");
                    return true;
                }
                Log($"❌ {companyName}: Tải {fileType} thất bại!");
                return false;
            }
            catch (Exception ex)
            {
                Log($"❌ {companyName}: Lỗi XulyexelraAsync: {ex.Message}");
                return false;
            }
        }

        #endregion

        #region Tải nhà cung cấp Viettel

        private async Task Tainhacungcap(string vbdbpath, DataRow companyRow = null)
        {
            string conn2 = AppConstants.BuildConnectionString(vbdbpath);
            var dtInvoiceInfo = ExecuteQuery2("SELECT * FROM tbInvoiceInfo", conn2);
            if (dtInvoiceInfo.Rows.Count == 0) return;

            var row = dtInvoiceInfo.Rows[0];
            string username = row["Username"]?.ToString();
            string password = row["Password"]?.ToString();

            using (HttpClientHandler handler = new HttpClientHandler())
            {
                handler.UseCookies = true;
                handler.CookieContainer = new CookieContainer();

                using (HttpClient client = new HttpClient(handler))
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                    client.DefaultRequestHeaders.Add("Accept", "application/json");

                    var json = $@"{{""username"": ""{username}"",""password"": ""{password}"",""rememberMe"": false,""captcha"": """"}}";
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(AppConstants.ViettelLoginUrl, content);
                    if (response.StatusCode == HttpStatusCode.Unauthorized)
                    {
                        Application.DoEvents();
                        await Task.Delay(3000);
                        Application.Exit();
                        return;
                    }

                    var cookies = handler.CookieContainer.GetCookies(new Uri(AppConstants.ViettelLoginUrl));
                    foreach (System.Net.Cookie cookie in cookies)
                    {
                        if (cookie.Name == "__cf_bm") useCookie.__cf_bm = cookie.Value;
                        if (cookie.Name == "JSESSIONID") useCookie.JSESSIONID = cookie.Value;
                        if (cookie.Name == "access_token") useCookie.access_token = cookie.Value;
                        if (cookie.Name == "session_token") useCookie.session_token = cookie.Value;
                    }

                    var result = await response.Content.ReadAsStringAsync();
                    loginResponse = JsonConvert.DeserializeObject<LoginResponse>(result);

                    if (loginResponse != null)
                    {
                        Application.DoEvents();
                        try
                        {
                            await ProcessViettelInvoices(client, conn2);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Lỗi xử lý: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private async Task ProcessViettelInvoices(HttpClient client, string conn2)
        {
            int page = 0, size = 1000, supplierId = 103807;

            int selectedMonth = int.Parse(comboBoxEdit1.EditValue.ToString());
            DateTime firstDayOfMonth = new DateTime(DateTime.Now.Year, selectedMonth, 1);

            string fromDate = firstDayOfMonth.AddHours(-7).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            string toDate = firstDayOfMonth.AddDays(1).AddSeconds(-1).AddHours(-7).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

            string urls = $"https://vinvoice.viettel.vn/api/cluster7/services/einvoiceapplication/api/invoice/search" +
                         $"?page={page}&size={size}" +
                         $"&createdDate.greaterThanOrEqual={Uri.EscapeDataString(fromDate)}" +
                         $"&createdDate.lessThanOrEqual={Uri.EscapeDataString(toDate)}" +
                         $"&supplierId.equals={supplierId}&dateType.equals=0" +
                         $"&invoiceStatus.equals=1&invoiceTypeId.notEquals=52" +
                         $"&sort=issueDate,desc&sort=invoiceNumber,desc";

            if (!string.IsNullOrEmpty(useCookie.access_token))
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", useCookie.access_token);

            var responses = await client.GetAsync(urls);

            if (!responses.IsSuccessStatusCode)
            {
                string error = await responses.Content.ReadAsStringAsync();
                MessageBox.Show($"Lỗi tìm kiếm: {responses.StatusCode}\n{error}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string jsonResult = await responses.Content.ReadAsStringAsync();
            var searchResult = JsonConvert.DeserializeObject<SearchResponse>(jsonResult);

            if (searchResult?.code != 200 || searchResult.data == null)
            {
                MessageBox.Show("Không tìm thấy hóa đơn nào trong khoảng thời gian này!",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var invoices = searchResult.data.content;
            var getResgistry = ExecuteQuery2("SELECT * FROM tbRegister", conn2);
            string savedPath = getResgistry.Rows[0]["Hoadonpath"].ToString();

            string pathYear = $"{AppConstants.YearPrefix}{firstDayOfMonth.Year}";
            int total = invoices.Count;

            for (int i = 0; i < invoices.Count; i++)
            {
                var invoice = invoices[i];
                string baseName = $"{invoice.createdDate:yyyyMMdd}_{mstcongty}_{invoice.invoiceNumber}_{invoice.invoiceSeri}";
                string filepath = Path.Combine(savedPath, pathYear, AppConstants.HdRaFolder, firstDayOfMonth.Month.ToString(), baseName + ".xml");
                string filepathpdf = Path.Combine(savedPath, pathYear, AppConstants.HdRaFolder, firstDayOfMonth.Month.ToString(), baseName + ".pdf");

                await DownloadInvoiceXml(client, invoice.id.ToString(), filepath);
                await DownloadInvoicePdf(client, invoice.id.ToString(), filepathpdf);

                await Task.Delay(150);
            }
        }

        private async Task DownloadInvoicePdf(HttpClient client, string invoiceId, string filePath)
        {
            try
            {
                string url = $"https://vinvoice.viettel.vn/api/cluster7/services/einvoicequery/api/invoice/gen-pdf?id={invoiceId}";
                Application.DoEvents();

                if (!string.IsNullOrEmpty(useCookie.access_token))
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", useCookie.access_token);

                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/pdf"));

                var response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    byte[] fileBytes = await response.Content.ReadAsByteArrayAsync();
                    Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                    File.WriteAllBytes(filePath, fileBytes);
                }
            }
            catch { }
        }

        private async Task DownloadInvoiceXml(HttpClient client, string invoiceId, string filePath)
        {
            try
            {
                string url = $"https://vinvoice.viettel.vn/api/cluster7/services/einvoiceapplication/api/invoice/downloadInvoiceFileXmlById?invoiceId={invoiceId}";
                Application.DoEvents();

                if (!string.IsNullOrEmpty(useCookie.access_token))
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", useCookie.access_token);

                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));

                var response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    string xmlContent = await response.Content.ReadAsStringAsync();
                    Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                    File.WriteAllText(filePath, xmlContent);
                }
            }
            catch { }
        }

        #endregion

        #region Get List Invoices

        private async Task<List<InvoiceInfo>> GetListHoaDonCanTai(string mstcongty, string savedPath, int type)
        {
            return await Task.Run(() =>
            {
                var result = new List<InvoiceInfo>();

                try
                {
                    string folderName = type == 1 ? AppConstants.HdVaoFolder : AppConstants.HdRaFolder;
                    string currentYear = $"{AppConstants.YearPrefix}{DateTime.Now.Year}";
                    string directoryPath = Path.Combine(savedPath, currentYear, folderName, comboBoxEdit1.EditValue.ToString());

                    if (!Directory.Exists(directoryPath))
                    {
                        Log($"❌ Thư mục không tồn tại: {directoryPath}");
                        return result;
                    }

                    var excelFiles = Directory.EnumerateFiles(directoryPath, "*.xlsx", SearchOption.AllDirectories)
                                              .Where(m => m.Contains(mstcongty)).ToList();

                    if (excelFiles.Count == 0)
                    {
                        Log($"📭 Không tìm thấy file Excel {folderName}");
                        return result;
                    }

                    int fileIndex = 1;
                    foreach (var excelFile in excelFiles)
                    {
                        using (var workbook = new XLWorkbook(excelFile))
                        {
                            var worksheet = workbook.Worksheet(1);
                            foreach (var row in worksheet.RowsUsed().Skip(3))
                            {
                                try
                                {
                                    string khhd = row.Cell("B").Value.ToString();
                                    string getSHHD = row.Cell("C").Value.ToString();
                                    string getSohd = RemoveLeadingZeros(row.Cell("D").Value.ToString());
                                    string GetNLap = row.Cell("E").Value.ToString();
                                    string mstnb = row.Cell("F").Value.ToString();

                                    if (!DateTime.TryParse(GetNLap, out DateTime getdate)) continue;
                                    if (getdate < dateEdit1.DateTime.Date || getdate > dateEdit2.DateTime.Date) continue;

                                    bool daTonTai = lookupHoaDonCT.Contains((mstnb, getSohd, getSHHD, getdate.Date, type));
                                    bool daTonTaiImport = lookupTbImport.Contains((mstnb, getSohd, getdate.Date, type));

                                    if (daTonTai || daTonTaiImport) continue;

                                    string filename = $"{getdate:yyyyMMdd}_{mstnb}_{getSohd}_{getSHHD}.xml";
                                    if (File.Exists(Path.Combine(directoryPath, filename))) continue;

                                    string url;
                                    if (type == 1)
                                    {
                                        if (fileIndex == 1 || fileIndex == 2)
                                            url = $"{AppConstants.GdtBaseUrl}/api/query/invoices/export-xml?nbmst={mstnb}&khhdon={getSHHD}&shdon={getSohd}&khmshdon={khhd}";
                                        else
                                            url = $"{AppConstants.GdtBaseUrl}/api/sco-query/invoices/export-xml?nbmst={mstnb}&khhdon={getSHHD}&shdon={getSohd}&khmshdon={khhd}";
                                    }
                                    else
                                    {
                                        if (fileIndex == 1)
                                            url = $"{AppConstants.GdtBaseUrl}/api/sco-query/invoices/export-xml?nbmst={mstnb}&khhdon={getSHHD}&shdon={getSohd}&khmshdon={khhd}";
                                        else
                                            url = $"{AppConstants.GdtBaseUrl}/api/query/invoices/export-xml?nbmst={mstnb}&khhdon={getSHHD}&shdon={getSohd}&khmshdon={khhd}";
                                    }

                                    result.Add(new InvoiceInfo
                                    {
                                        Mst = mstnb,
                                        SHHD = getSHHD,
                                        Sohd = getSohd,
                                        NLap = getdate,
                                        Khhd = khhd,
                                        DirectoryPath = directoryPath,
                                        Url = url,
                                        Type = type
                                    });

                                    if (result.Count % 100 == 0)
                                        Log($"📋 Đã đọc {result.Count} hóa đơn cần tải...");
                                }
                                catch (Exception ex) { Log($"Lỗi dòng: {ex.Message}"); }
                            }
                        }
                        fileIndex++;
                    }
                }
                catch (Exception ex) { Log($"❌ Lỗi GetListHoaDonCanTai: {ex.Message}"); }

                return result;
            });
        }

        #endregion

        #region Download Invoices

        private async Task TaiHangLoatHoaDon(CompanyHttpClient client, List<InvoiceInfo> invoices,
     string typeName, string companyName, DataRow companyRow = null)
        {
            if (invoices == null || invoices.Count == 0)
            {
                Log(string.Format("📭 {0}: Không có hóa đơn {1} để tải", companyName, typeName));
                return;
            }

            Log(string.Format("📥 {0}: Bắt đầu tải {1} hóa đơn {2}...", companyName, invoices.Count, typeName));

            int total = invoices.Count;
            var stats = new DownloadStats();
            var timeoutSeconds = int.Parse(txttimeout.Text);
            int maxParallel = 3;
            var sw = Stopwatch.StartNew();

            var semaphore = new SemaphoreSlim(maxParallel);
            var tasks = new List<Task>();

            foreach (var invoice in invoices)
            {
                var capturedInvoice = invoice; // tránh closure bug
                tasks.Add(Task.Run(async () =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        using (var cts = new CancellationTokenSource())
                        {
                            cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));

                            bool ok = await DownloadSingleInvoiceAsync(client, capturedInvoice, companyName, cts.Token);

                            lock (stats.LockObj)
                            {
                                if (ok) stats.Downloaded++;
                                else stats.Failed++;

                                int done = stats.Downloaded + stats.Failed;
                                if (done % 5 == 0 || done == total)
                                {
                                    UpdateStatusOnUI(companyRow,
                                        string.Format("📥 {0}: {1}/{2} (✅{3})", typeName, done, total, stats.Downloaded));
                                }
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        lock (stats.LockObj) { stats.TimeoutCount++; stats.Failed++; }
                        Log(string.Format("⏰ {0}: Timeout HĐ {1}", companyName, capturedInvoice.Sohd));
                    }
                    catch (Exception ex)
                    {
                        lock (stats.LockObj) stats.Failed++;
                        Log(string.Format("❌ {0}: Lỗi HĐ {1}: {2}", companyName, capturedInvoice.Sohd, ex.Message));
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }));
            }

            await Task.WhenAll(tasks);
            semaphore.Dispose();

            sw.Stop();
            string finalStatus = string.Format("✅ Đã tải {0}: {1}/{2} (✅{3} ❌{4} ⏰{5}) - {6:F1}s",
                typeName, stats.Downloaded, total, stats.Downloaded, stats.Failed, stats.TimeoutCount,
                sw.Elapsed.TotalSeconds);
            Log(string.Format("✅ {0}: {1}", companyName, finalStatus));

            if (companyRow != null) UpdateStatusOnUI(companyRow, finalStatus);
        }

        private async Task<bool> DownloadSingleInvoiceAsync(CompanyHttpClient client, InvoiceInfo invoice,
            string companyName, CancellationToken ct)
        {
            if (invoice == null) return false;

            string filename = $"{invoice.NLap:yyyyMMdd}_{invoice.Mst}_{invoice.Sohd}_{invoice.SHHD}.zip";
            string path = Path.Combine(invoice.DirectoryPath, filename);
            string pathXml = Path.ChangeExtension(path, ".xml");

            if (File.Exists(path) || File.Exists(pathXml)) return true;

            for (int retry = 1; retry <= AppConstants.MaxRetryAttempts; retry++)
            {
                ct.ThrowIfCancellationRequested();

                try
                {
                    bool ok = await client.DownloadFileAsync(invoice.Url, path, int.Parse(txttimeout.Text), ct);
                    if (ok)
                    {
                        ExtractZipXML(path);
                        Log($"✅ {companyName}: Tải HĐ {invoice.Sohd} thành công");
                        Taihoadongoc(pathXml);
                        return true;
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    Log($"⚠️ {companyName}: Lỗi HĐ {invoice.Sohd} lần {retry}: {ex.Message}");
                }

                if (retry < AppConstants.MaxRetryAttempts)
                    await Task.Delay(500 * retry, ct);
            }

            // Fallback KNM
            if (invoice.Type == 1)
            {
                await GetKNMXMLAsync(invoice.Mst, invoice.SHHD, invoice.Sohd, tokken,
                    invoice.NLap, invoice.DirectoryPath, invoice.Sohd, companyName);
            }

            return false;
        }
        public async Task Taihoadongoc(string pathXml)
        {
            try
            {
                if (string.IsNullOrEmpty(pathXml) || !File.Exists(pathXml))
                {
                    Console.WriteLine($"❌ [Taihoadongoc] File không tồn tại: {pathXml}");
                    return;
                }

                var xmlDoc = new XmlDocument();
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Ignore,
                    XmlResolver = null,
                    CheckCharacters = false,
                    IgnoreWhitespace = true
                };

                using (var xmlReader = XmlReader.Create(pathXml, settings))
                    xmlDoc.Load(xmlReader);

                XmlNode root = xmlDoc.DocumentElement;
                if (root == null) return;

                XmlNode ttChung = root.SelectSingleNode("//TTChung");
                XmlNode nBan = root.SelectSingleNode("//NBan");
                if (ttChung == null) return;

                string SHDon = Helpers.RemoveLeadingZeros(ttChung.SelectSingleNode("SHDon")?.InnerText ?? "");
                string nccRaw = ttChung.SelectSingleNode("MSTTCGP")?.InnerText;
                string NCC = GetNCCNameCached(nccRaw);
                string mabimat = GetSecretCode(pathXml, NCC);

                int type = pathXml.IndexOf("HDVao", StringComparison.OrdinalIgnoreCase) >= 0 ? 1 : 2;
                string mst = nBan?.SelectSingleNode("MST")?.InnerText ?? "";

                string rootFolder = GetRootFolderFromXmlPath(pathXml);
                if (string.IsNullOrEmpty(rootFolder))
                {
                    Console.WriteLine("❌ [Taihoadongoc] Không xác định được thư mục gốc");
                    return;
                }

                // ✅ CHỈ ENQUEUE — KHÔNG MỞ CHROME
                if (NCC == "VIETTEL" || NCC == "MISA" || NCC == "FAST" || NCC == "BKAV")
                {
                    var task = new InvoiceDownloadTask
                    {
                        PathXml = pathXml,
                        Mst = mst,
                        Shdon = SHDon,
                        Ncc = NCC,
                        Mabimat = mabimat,
                        Type = type,
                        RootFolder = rootFolder
                    };

                    _pdfQueue.Enqueue(task);
                    _pdfSignal.Release();
                    Log($"📥 [Queue] +1 HĐ {SHDon} ({NCC}) — còn {_pdfQueue.Count} chờ");
                }
                else
                {
                    Log($"ℹ️ [Taihoadongoc] Bỏ qua NCC {NCC} cho HĐ {SHDon}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ [Taihoadongoc] Lỗi: {ex.Message}");
            }
        }
        private readonly ConcurrentDictionary<string, string> _nccCache
= new ConcurrentDictionary<string, string>();
        private string GetNCCNameCached(string code)
        {
            if (string.IsNullOrEmpty(code)) return "";
            return _nccCache.GetOrAdd(code, c => GetNCCName(c));
        }
        public string GetNCCName(string code)
        {
            if (code == null)
                return "";
            if (code == "0100109106")
            {
                return "VIETTEL";
            }
            if (code == "0101243150")
            {
                return "MISA";
            }
            if (code == "0101360697")
            {
                return "BKAV";
            }
            if (code == "0100727825")
            {
                return "FAST";
            }
            if (code.Contains("0106026495-001"))
            {
                return "MINVOICE";
            }
            if (code.Contains("0106026495"))
            {
                return "MINVOICE";
            }
            if (code == "0312303803")
            {
                return "Win Tech";
            }
            if (code == "0100684378")
            {
                return "VNPT";
            }
            if (code == "0105232093")
            {
                return "Cyber";
            }
            if (code == "0105987432")
            {
                return "Softdream";
            }
            if (code == "0401486901")
            {
                return "Visnam";
            }
            return "";
        }
        private void StartPdfWorker()
        {
            if (_pdfWorkerRunning) return;
            _pdfWorkerRunning = true;
            _pdfCts = new CancellationTokenSource();
            _pdfWorkerTask = Task.Run(() => PdfWorkerLoop(_pdfCts.Token));
            Log("🚀 [PdfWorker] Đã khởi động");
        }

        private async Task StopPdfWorkerAsync()
        {
            if (!_pdfWorkerRunning) return;
            _pdfCts?.Cancel();
            _pdfSignal.Release(1000); // đánh thức worker
            if (_pdfWorkerTask != null)
            {
                try { await _pdfWorkerTask; } catch { }
            }
            _pdfWorkerRunning = false;
            Log("🛑 [PdfWorker] Đã dừng");
        }

        private async Task PdfWorkerLoop(CancellationToken ct)
        {
            var semaphore = new SemaphoreSlim(MAX_PARALLEL_PDF, MAX_PARALLEL_PDF);
            var runningTasks = new List<Task>();
            int processedInBatch = 0;

            while (!ct.IsCancellationRequested)
            {
                try
                {
                    await _pdfSignal.WaitAsync(ct);
                    if (ct.IsCancellationRequested) break;

                    if (!_pdfQueue.TryDequeue(out var task))
                        continue;

                    // Chờ slot trống
                    await semaphore.WaitAsync(ct);

                    // Chạy trên thread riêng
                    var t = Task.Run(async () =>
                    {
                        try
                        {
                            Log($"⚙️ [PdfWorker] Bắt đầu HĐ {task.Shdon} ({task.Ncc})");
                            await ProcessDownloadTaskAsync(task, ct);
                            Log($"✅ [PdfWorker] Xong HĐ {task.Shdon}");
                        }
                        catch (Exception ex)
                        {
                            Log($"❌ [PdfWorker] Lỗi HĐ {task.Shdon}: {ex.Message}");
                        }
                        finally
                        {
                            semaphore.Release();
                        }
                    }, ct);

                    runningTasks.Add(t);

                    // Dọn task đã xong
                    runningTasks.RemoveAll(x => x.IsCompleted);

                    processedInBatch++;

                    // Nghỉ sau mỗi batch
                    if (processedInBatch >= PDF_BATCH_SIZE)
                    {
                        Log($"⏸ [PdfWorker] Nghỉ {PDF_BATCH_DELAY_MS}ms (còn {_pdfQueue.Count} chờ)");
                        processedInBatch = 0;
                        await Task.Delay(PDF_BATCH_DELAY_MS, ct);
                    }
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    Log($"❌ [PdfWorker] Lỗi: {ex.Message}");
                    await Task.Delay(1000, ct);
                }
            }

            // Chờ tất cả task xong
            try { await Task.WhenAll(runningTasks); } catch { }
            semaphore.Dispose();
        }

        /// <summary>
        /// Xử lý 1 task tải PDF — GỌI LOGIC CŨ (TraCuuNCC / GetHoaDonMisa)
        /// </summary>
        private async Task ProcessDownloadTaskAsync(InvoiceDownloadTask task, CancellationToken ct)
        {
            var dir = new DirectoryInfo(Path.GetDirectoryName(task.PathXml));
            while (dir != null && !string.Equals(dir.Name, "Hoadon", StringComparison.OrdinalIgnoreCase))
                dir = dir.Parent;

            string savedPath = dir?.FullName ?? task.RootFolder;
            if (task.Ncc == "MISA")
            {
               // await GetHoaDonMisa(true, task.Mabimat, task.Mst, task.Shdon, task.Type, savedPath);
            }
            else if (task.Ncc == "VIETTEL")
            {
               // await TraCuuNCC(true, task.Mst, task.Mabimat, task.Shdon, savedPath, task.Type);
            }
            else if (task.Ncc == "FAST")
            {
                await TaiFAST(true, task.Mst, task.Shdon,task.Khhd, task.Mabimat, savedPath, task.Type);
            }
            else if (task.Ncc == "BKAV")
            {
                await TaiBKAV(true, task.Mst, task.Shdon, task.Khhd, task.Mabimat, savedPath, task.Type);
            }
        }
        public async Task TaiBKAV(bool toggle, string mst, string shdon, string KHHDon, string mabimat, string savedPath, int type)
        {
            string url = $"https://van.ehoadon.vn/Lookup?InvoiceGUID={mabimat}";
            string downloadPath = "";
            await Task.Run(async () =>
            {
                if (type==1)
                    downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDVao\\" + dateEdit1.DateTime.Month;
                if (type==2)
                    downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\" + "\\HDRa\\" + dateEdit1.DateTime.Month;

                Directory.CreateDirectory(downloadPath);

                ChromeOptions options = new ChromeOptions();
                //options.AddArgument("--headless=new");
                options.AddArgument("--disable-gpu");
                options.AddArgument("--window-size=1920,1080");
                options.AddUserProfilePreference(
                    "download.default_directory",
                    downloadPath);

                options.AddUserProfilePreference(
                    "download.prompt_for_download",
                    false);

                options.AddUserProfilePreference(
                    "download.directory_upgrade",
                    true);

                options.AddUserProfilePreference(
                    "safebrowsing.enabled",
                    true);


                string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                service.HideCommandPromptWindow = true; // Ẩn cửa sổ cmd đen

               var driver = new ChromeDriver(service, options);

                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                driver.Navigate().GoToUrl(url);

               
                // Chờ trang load xong
                wait.Until(d =>
                {
                    try
                    {
                        return ((IJavaScriptExecutor)d)
                            .ExecuteScript("return document.readyState")
                            .ToString() == "complete";
                    }
                    catch
                    {
                        return false;
                    }
                });

                // ==========================================
                // 1. Tìm nút Download
                // ==========================================

                IWebElement btnDownload = wait.Until(d =>
                {
                    try
                    {
                        IWebElement el = d.FindElement(By.Id("btnDownload"));

                        return el.Displayed && el.Enabled
                            ? el
                            : null;
                    }
                    catch
                    {
                        return null;
                    }
                });

                // ==========================================
                // 2. Hover vào Download
                // ==========================================

                Actions actions = new Actions(driver);

                actions.MoveToElement(btnDownload).Perform();

                // ==========================================
                // 3. Chờ menu dropdown hiện
                // ==========================================

                wait.Until(d =>
                {
                    try
                    {
                        IWebElement menu = d.FindElement(
                            By.CssSelector("#divDownloads ul.dropdown-menu"));

                        return menu.Displayed;
                    }
                    catch
                    {
                        return false;
                    }
                });

                // ==========================================
                // 4. Tìm Download PDF
                // ==========================================

                IWebElement pdfButton = wait.Until(d =>
                {
                    try
                    {
                        IWebElement el = d.FindElement(
                            By.Id("LinkDownPDF"));

                        return el.Displayed && el.Enabled
                            ? el
                            : null;
                    }
                    catch
                    {
                        return null;
                    }
                });

                // ==========================================
                // 5. Click Download PDF
                // ==========================================

                pdfButton.Click();

                // ==========================================
                // 6. Chờ file tải xong
                // ==========================================

                string pdfFile = WaitForPdf(shdon, mst, downloadPath, 10);
                if (pdfFile == null)
                {
                    XtraMessageBox.Show("Không thể tải");
                    driver.Quit();
                }
                Thread.Sleep(500);
                if (pdfFile != null)
                {
                    string fileMoi = Path.Combine(
                        downloadPath,
                        $"{mst}_{shdon}_{KHHDon}.pdf"
                    );

                    // Nếu tên mới đã tồn tại thì xóa
                    if (File.Exists(fileMoi))
                        File.Delete(fileMoi);

                    // Đổi tên file PDF vừa tải
                    File.Move(pdfFile, fileMoi);
                    if (toggle)
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = fileMoi,
                            UseShellExecute = true
                        });
                        Console.WriteLine($"📂 Đã mở file: {fileMoi}");
                    }
                    Console.WriteLine($"Đã đổi tên PDF: {fileMoi}");
                    driver.Quit();
                }
            });
        }
        public async Task TaiFAST(bool toggle, string mst, string shdon, string KHHDon, string mabimat,string savedPath,int type)
        {
            if (KHHDon == null)
                KHHDon = "";
            string url = "https://invoice.fast.com.vn/tra-cuu-hoa-don-dien-tu/";
            string downloadPath = "";

            if (type==1)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDVao\\" + dateEdit1.DateTime.Month;
            if (type == 2)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDRa\\" + dateEdit1.DateTime.Month;

            Directory.CreateDirectory(downloadPath);
            await Task.Run(async () =>
            {
                ChromeOptions options = new ChromeOptions();
                //options.AddArgument("--headless=new");
                options.AddArgument("--disable-gpu");
                options.AddArgument("--window-size=1920,1080");
                options.AddArgument("--disable-notifications");
                options.AddArgument("--disable-infobars");
                options.AddArgument("--no-first-run");
                options.AddArgument("--no-default-browser-check");
                options.AddArgument("--disable-blink-features=AutomationControlled");

                // Tắt PDF Viewer → nếu có redirect PDF thì Chrome tự tải (dự phòng)
                options.AddUserProfilePreference("plugins.always_open_pdf_externally", true);

                // Download
                options.AddUserProfilePreference("download.default_directory", downloadPath);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("safebrowsing.enabled", true);
                options.AddUserProfilePreference("profile.default_content_setting_values.popups", 2);

                // Load uBlock Origin nếu có
                string ublockPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Extensions", "uBlock");
                if (Directory.Exists(ublockPath))
                    options.AddArgument("--load-extension=" + ublockPath);

                IWebDriver driver = null;
                try
                {
                    // [SỬA QUAN TRỌNG]: Ép ChromeDriverService trỏ thẳng vào thư mục App để bỏ qua SeleniumManager đang bị lỗi path
                    string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                    service.HideCommandPromptWindow = true; // Ẩn cửa sổ cmd đen

                    driver = new ChromeDriver(service, options);

                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                    driver.Navigate().GoToUrl(url);
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

                    // ============================================================
                    // 1. Điền mã vào ô keyword
                    // ============================================================
                    var keywordInput = wait.Until(d =>
                    {
                        try
                        {
                            var el = d.FindElement(By.CssSelector("div.btn-search-invoice input[name='keyword']"));
                            return (el != null && el.Displayed) ? el : null;
                        }
                        catch (NoSuchElementException) { return null; }
                    });

                    keywordInput.Click();
                    keywordInput.Clear();
                    keywordInput.SendKeys(mabimat);

                    // Trigger event cho React/Angular
                    js.ExecuteScript(@"
            var el = arguments[0];
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        ", keywordInput);

                    // Lưu window chính
                    string mainWindow = driver.CurrentWindowHandle;

                    // ============================================================
                    // 2. Bấm nút search
                    // ============================================================
                    var btnSearch = wait.Until(d =>
                    {
                        try
                        {
                            var btn = d.FindElement(By.CssSelector("div.btn-search-invoice button.button-invoice"));
                            return (btn != null && btn.Displayed && btn.Enabled) ? btn : null;
                        }
                        catch (NoSuchElementException) { return null; }
                    });

                    js.ExecuteScript("arguments[0].click();", btnSearch);

                    // ============================================================


                    // ============================================================
                    // 4. Chờ Fast redirect sang URL PDF (einvoice.fast.com.vn)
                    // ============================================================
                    Thread.Sleep(1000);
                    string pdfFile = WaitForPdf(shdon, mst, downloadPath, 10);
                    if (pdfFile == null)
                    {
                        XtraMessageBox.Show("Không thể tải");
                        driver.Quit();
                    }
                    if (pdfFile != null && File.Exists(pdfFile))
                    {
                        string fileMoi = Path.Combine(
                            downloadPath,
                            mst + "_" + shdon + "_" + KHHDon + ".pdf"
                        );

                        // Nếu file nguồn trùng tên đích → bỏ qua
                        if (string.Equals(pdfFile, fileMoi, StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine("⚠ File nguồn đã đúng tên đích, bỏ qua Move.");
                        }
                        else
                        {
                            if (File.Exists(fileMoi))
                                File.Delete(fileMoi);

                            try
                            {
                                File.Move(pdfFile, fileMoi);
                                Console.WriteLine("✅ Đã đổi tên PDF: " + fileMoi);
                            }
                            catch (FileNotFoundException)
                            {
                                Console.WriteLine("⚠ File nguồn biến mất trước khi Move: " + pdfFile);
                                // Thử tìm lại file PDF mới nhất

                                if (pdfFile != null && File.Exists(pdfFile))
                                    File.Move(pdfFile, fileMoi);
                            }
                            catch (IOException ioEx)
                            {
                                Console.WriteLine("⚠ Lỗi IO khi Move: " + ioEx.Message);
                            }
                        }

                        if (toggle && File.Exists(fileMoi))
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = fileMoi,
                                UseShellExecute = true
                            });
                            Console.WriteLine("📂 Đã mở file: " + fileMoi);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("❌ Lỗi TaiFAST: " + ex.Message);
                    throw;
                }
                finally
                {
                    try { if (driver != null) driver.Quit(); } catch { }
                    try { if (driver != null) driver.Dispose(); } catch { }

                    // Kill chromedriver sót
                    try
                    {
                        foreach (var p in System.Diagnostics.Process.GetProcessesByName("chromedriver"))
                        {
                            try { p.Kill(); } catch { }
                        }
                    }
                    catch { }
                }
            });
        }
        private string WaitForPdf(string shdon, string mst, string folder, int timeoutSeconds)
        {
            DateTime start = DateTime.Now;

            while ((DateTime.Now - start).TotalSeconds < timeoutSeconds)
            {
                try
                {
                    // Còn file đang tải dở → chờ tiếp
                    string[] downloading = Directory.GetFiles(folder, "*.crdownload");
                    if (downloading.Length > 0)
                    {
                        Thread.Sleep(500);
                        continue;
                    }

                    // Lấy tất cả file PDF
                    string[] pdfFiles = Directory.GetFiles(folder, "*.pdf");

                    // Lọc chỉ những file có tên chứa shdon
                    var matched = pdfFiles
      .Where(f =>
      {
          string fileName = Path.GetFileName(f);

          return fileName.IndexOf(shdon, StringComparison.OrdinalIgnoreCase) >= 0
              || fileName.IndexOf(mst, StringComparison.OrdinalIgnoreCase) >= 0;
      })
      .ToArray();
                    if (matched.Length > 0)
                    {
                        // Chọn file mới nhất trong số các file khớp
                        string newestFile = null;
                        DateTime newestTime = DateTime.MinValue;

                        foreach (string file in matched)
                        {
                            DateTime time = File.GetLastWriteTime(file);
                            if (time > newestTime)
                            {
                                newestTime = time;
                                newestFile = file;
                            }
                        }

                        // Đảm bảo file đã ghi xong (size ổn định)
                        long size1 = new FileInfo(newestFile).Length;
                        Thread.Sleep(300);
                        long size2 = new FileInfo(newestFile).Length;

                        if (size1 == size2 && size1 > 0)
                            return newestFile;

                        // File đang ghi tiếp → chờ
                        Thread.Sleep(500);
                        continue;
                    }
                }
                catch (IOException)
                {
                    // File đang bị lock hoặc lỗi IO tạm thời → chờ rồi thử lại
                    Thread.Sleep(500);
                }
                catch (UnauthorizedAccessException)
                {
                    Thread.Sleep(500);
                }

                Thread.Sleep(500);
            }

            // Hết timeout mà không tìm được file
            Console.WriteLine($"⏱ Hết {timeoutSeconds}s không tìm thấy PDF chứa '{shdon}' trong: {folder}");
            return null;
        }
        private async Task GetHoaDonMisa(bool toggle, string maBiMat, string mst, string shdon, int type, string savedPath)
        {
            // ===== 1. XÁC ĐỊNH THƯ MỤC ĐÍCH (dùng chung cho cả download và lưu file) =====
            string year = dateEdit1.DateTime.Year.ToString();
            string month = dateEdit1.DateTime.Month.ToString();

            await Task.Run(async () =>
            {
                string baseDir;
                if (type == 1)
                    baseDir = Path.Combine(savedPath, $"HD{year}", "HDVao", month);
                else if (type == 2)
                    baseDir = Path.Combine(savedPath, $"HD{year}", "HDRa", month);
                else
                    baseDir = Path.Combine(savedPath, $"HD{year}", "HDKhac", month);

                Directory.CreateDirectory(baseDir);

                // Thư mục tạm để Chrome tải file về (tách riêng khỏi thư mục lưu chính thức)
                string tempDownloadFolder = Path.Combine(Path.GetTempPath(), "MisaTemp", Guid.NewGuid().ToString("N"));

                // 2. Tạo thư mục thực tế trên ổ đĩa
                Directory.CreateDirectory(tempDownloadFolder);

                // [SỬA QUAN TRỌNG]: Chrome trên .NET Framework yêu cầu đổi dấu \ thành / để tránh lỗi "path is not of a legal form"
                string chromeDownloadPath = tempDownloadFolder.Replace("\\", "/");

                // Cấu hình ChromeOptions
                // Cấu hình ChromeOptions
                ChromeOptions options = new ChromeOptions();

                // === CHẠY ẨN (HEADLESS) & CHỐNG CHẶN TẢI FILE ===
                //  options.AddArgument("--headless");
                options.AddArgument("--disable-gpu");
                options.AddArgument("--no-sandbox");
                options.AddArgument("--disable-dev-shm-usage");
                options.AddArgument("--window-size=1920,1080");
                options.AddArgument("--disable-extensions");
                options.AddArgument("--disable-popup-blocking");
                options.AddArgument("--disable-web-security");

                // Cấu hình tải file
                options.AddUserProfilePreference("download.default_directory", tempDownloadFolder);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("safebrowsing.enabled", true);
                options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);

                IWebDriver driver = null;
                try
                {
                    // [SỬA QUAN TRỌNG]: Ép ChromeDriverService trỏ thẳng vào thư mục App để bỏ qua SeleniumManager đang bị lỗi path
                    string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                    service.HideCommandPromptWindow = true; // Ẩn cửa sổ cmd đen

                    driver = new ChromeDriver(service, options);
                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                    // 1. Điều hướng đến trang tra cứu
                    Console.WriteLine("Đang truy cập trang tra cứu...");
                    driver.Navigate().GoToUrl("https://www.meinvoice.vn/tra-cuu/");

                    // 2. Nhập mã bí mật
                    IWebElement inputCode = wait.Until(drv => drv.FindElement(By.Id("txtCode")));
                    inputCode.Clear();
                    inputCode.SendKeys(maBiMat);
                    Console.WriteLine($"Đã nhập mã: {maBiMat}");

                    // 3. Nhấn Enter
                    inputCode.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Console.WriteLine("Đã gửi yêu cầu tra cứu.");

                    // 4. Đợi popup hiển thị
                    Thread.Sleep(500);

                    // 5. Kiểm tra popup có xuất hiện không
                    IWebElement popupDiv = null;
                    try
                    {
                        popupDiv = wait.Until(drv => drv.FindElement(By.Id("showContentPopupInvoice")));
                        Console.WriteLine("Đã tìm thấy popup.");
                    }
                    catch (WebDriverTimeoutException)
                    {
                        Console.WriteLine("⚠️ Không tìm thấy popup hóa đơn.");
                    }

                    // 6. Click "Tải hóa đơn" (bước 1)
                    try
                    {
                        var downloadSpan = popupDiv.FindElement(By.CssSelector("span.download-invoice"));
                        downloadSpan.Click();
                        Console.WriteLine("Đã click vào nút 'Tải hóa đơn'.");
                    }
                    catch (NoSuchElementException)
                    {
                        Console.WriteLine("⚠️ Không tìm thấy nút 'Tải hóa đơn'.");
                    }

                    // 7. Đợi popup chọn định dạng xuất hiện
                    Thread.Sleep(500);

                    // ===== 8. CHỤP DANH SÁCH FILE TRƯỚC KHI CLICK PDF =====
                    var beforeFiles = new HashSet<string>(
                        Directory.GetFiles(tempDownloadFolder, "*.*")
                                 .Select(f => Path.GetFileName(f))
                    );

                    // 9. Click "Tải hóa đơn dạng PDF" (bước 2)
                    try
                    {
                        var pdfOption = wait.Until(drv => drv.FindElement(By.CssSelector("div.dm-item.pdf.txt-download-pdf")));
                        pdfOption.Click();
                        Console.WriteLine("Đã click vào 'Tải hóa đơn dạng PDF'.");
                    }
                    catch (WebDriverTimeoutException)
                    {
                        // Thử tìm bằng cách khác
                        try
                        {
                            var optionsList = driver.FindElements(By.CssSelector("div.dm-item"));
                            foreach (var opt in optionsList)
                            {
                                if (opt.Text.Contains("PDF") || opt.Text.Contains("pdf"))
                                {
                                    opt.Click();
                                    Console.WriteLine("Đã click vào option PDF (tìm bằng text).");
                                    break;
                                }
                            }
                        }
                        catch (Exception ex2)
                        {
                            Console.WriteLine($"⚠️ Không tìm thấy option PDF: {ex2.Message}");
                        }
                    }

                    // ===== 10. CHỜ FILE PDF MỚI XUẤT HIỆN VÀ TẢI XONG =====
                    Console.WriteLine("Đang chờ file tải về...");
                    string latestPdfFile = null;
                    var deadline = DateTime.Now.AddSeconds(60);
                    while (DateTime.Now < deadline)
                    {
                        // Lấy các file mới (không có trong danh sách beforeFiles)
                        var newFiles = Directory.GetFiles(tempDownloadFolder, "*.pdf")
                                                .Where(f => !beforeFiles.Contains(Path.GetFileName(f)))
                                                .ToList();

                        if (newFiles.Count > 0)
                        {
                            var candidate = newFiles.OrderByDescending(File.GetCreationTime).First();

                            // Kiểm tra file đã ghi xong chưa (không bị lock)
                            try
                            {
                                using (var fs = File.Open(candidate, FileMode.Open, FileAccess.Read, FileShare.None))
                                {
                                    fs.Close();
                                }
                                latestPdfFile = candidate;
                                break;
                            }
                            catch (IOException)
                            {
                                // File đang bị Chrome giữ, chờ tiếp
                            }
                        }
                        Thread.Sleep(500);
                    }

                    if (latestPdfFile == null)
                    {
                        Console.WriteLine("❌ Không tải được file PDF hoặc quá thời gian chờ!");
                        return;
                    }

                    Console.WriteLine($"📄 File tải về: {latestPdfFile}");

                    // ===== 11. TẠO ĐƯỜNG DẪN ĐÍCH, KHÔNG GHI ĐÈ =====
                    string safeMst = MakeSafeFileName(mst);
                    string safeShdon = MakeSafeFileName(shdon);
                    string newPath = Path.Combine(baseDir, $"{safeMst}_{safeShdon}.pdf");

                    // Nếu file đích đã tồn tại → thêm hậu tố _1, _2, ...
                    if (File.Exists(newPath))
                    {
                        string dir = Path.GetDirectoryName(newPath);
                        string nameNoExt = Path.GetFileNameWithoutExtension(newPath);
                        string ext = Path.GetExtension(newPath);
                        int counter = 1;
                        do
                        {
                            newPath = Path.Combine(dir, $"{nameNoExt}_{counter}{ext}");
                            counter++;
                        } while (File.Exists(newPath));

                        Console.WriteLine($"⚠️ File gốc đã tồn tại, lưu thành: {Path.GetFileName(newPath)}");
                    }

                    // ===== 12. MOVE FILE AN TOÀN =====
                    File.Move(latestPdfFile, newPath);
                    Console.WriteLine($"✅ Đã lưu: {newPath}");

                    // ===== 13. XÓA FILE RÁC CÒN LẠI TRONG THƯ MỤC TẠM =====
                    try
                    {
                        foreach (var f in Directory.GetFiles(tempDownloadFolder))
                        {
                            try { File.Delete(f); } catch { }
                        }
                    }
                    catch { }

                    // ===== 14. MỞ FILE NẾU toggle = true =====
                    if (toggle)
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = newPath,
                            UseShellExecute = true
                        });
                        Console.WriteLine($"📂 Đã mở file: {newPath}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Lỗi: {ex.Message}");
                }
                finally
                {
                    // Đợi Chrome ghi xong hoàn toàn trước khi đóng
                    Thread.Sleep(2000);

                    if (driver != null)
                    {
                        try { driver.Quit(); } catch { }
                        Console.WriteLine("Đã đóng trình duyệt.");
                    }

                    // Dọn thư mục tạm
                    try
                    {
                        if (Directory.Exists(tempDownloadFolder))
                            Directory.Delete(tempDownloadFolder, true);
                    }
                    catch { }
                }
            });
        }
        /// <summary>
        /// Từ pathXml → lấy thư mục gốc chứa hóa đơn
        /// Ví dụ: "D:\Hoadon\HD2026\HDVao\10\file.xml" → "D:\Hoadon"
        /// </summary>
        private static string GetRootFolderFromXmlPath(string pathXml)
        {
            try
            {
                if (string.IsNullOrEmpty(pathXml)) return "";

                // Lấy thư mục chứa file .xml: "D:\Hoadon\HD2026\HDVao\10"
                string dir = Path.GetDirectoryName(pathXml);
                if (string.IsNullOrEmpty(dir)) return "";

                // Đi lên 3 cấp
                string parent1 = Path.GetDirectoryName(dir);       // "D:\Hoadon\HD2026\HDVao"
                string parent2 = Path.GetDirectoryName(parent1);   // "D:\Hoadon\HD2026"
                string parent3 = Path.GetDirectoryName(parent2);   // "D:\Hoadon"

                // Ưu tiên cấp cao nhất tồn tại
                if (!string.IsNullOrEmpty(parent3) && Directory.Exists(parent3))
                    return parent3;

                if (!string.IsNullOrEmpty(parent2) && Directory.Exists(parent2))
                    return parent2;

                if (!string.IsNullOrEmpty(parent1) && Directory.Exists(parent1))
                    return parent1;

                return dir;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ GetRootFolderFromXmlPath: {ex.Message}");
                return Path.GetDirectoryName(pathXml) ?? "";
            }
        }
        static string SaveBase64ToFile(string dataUrl, string name)
        {
            try
            {
                if (string.IsNullOrEmpty(dataUrl) || !dataUrl.Contains(","))
                    return "";

                // ✅ Thư mục riêng của app, KHÔNG dùng Path.GetTempPath()
                string appDataDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ToolTaiHD", "Captcha");

                Directory.CreateDirectory(appDataDir);

                // ✅ Tên file ngắn gọn — không dùng GUID dài
                string fileName = $"{name}_{DateTime.Now:HHmmssfff}.png";
                string path = Path.Combine(appDataDir, fileName);

                var base64 = dataUrl.Substring(dataUrl.IndexOf(",") + 1);
                var bytes = Convert.FromBase64String(base64);

                Console.WriteLine($"[SaveBase64] Bytes = {bytes.Length}, Magic = " +
                    BitConverter.ToString(bytes, 0, Math.Min(8, bytes.Length)));

                File.WriteAllBytes(path, bytes);
                Console.WriteLine($"[SaveBase64] ✅ Đã lưu: {path}");

                return path;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SaveBase64] ❌ Lỗi: {ex.Message}");
                return "";
            }
        }
        static int FindHolePosition2(string imagePath)
        {
            try
            {
                if (!File.Exists(imagePath)) return -1;

                int holeX = -1;
                int minX = int.MaxValue;
                int maxX = int.MinValue;
                long totalDarkPixels = 0;

                // Sử dụng Bitmap thuần túy của .NET Framework
                using (Bitmap bmp = new Bitmap(imagePath))
                {
                    int width = bmp.Width;
                    int height = bmp.Height;

                    // Dùng LockBits để quét điểm ảnh cực nhanh (thay vì GetPixel chậm)
                    System.Drawing.Rectangle rect = new System.Drawing.Rectangle(0, 0, width, height);
                    BitmapData bmpData = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

                    int bytesPerPixel = 3;
                    int stride = bmpData.Stride;
                    IntPtr ptr = bmpData.Scan0;

                    int bytes = Math.Abs(stride) * height;
                    byte[] rgbValues = new byte[bytes];

                    System.Runtime.InteropServices.Marshal.Copy(ptr, rgbValues, 0, bytes);
                    bmp.UnlockBits(bmpData);

                    // Quét từng pixel để tìm vùng tối (tương đương ThresholdBinaryInv với ngưỡng 30)
                    for (int y = 0; y < height; y++)
                    {
                        int rowOffset = y * stride;
                        for (int x = 0; x < width; x++)
                        {
                            int idx = rowOffset + x * bytesPerPixel;
                            byte b = rgbValues[idx];
                            byte g = rgbValues[idx + 1];
                            byte r = rgbValues[idx + 2];

                            // Tính độ sáng trung bình (Grayscale)
                            int gray = (r + g + b) / 3;

                            // Nếu là vùng tối (ngưỡng < 30)
                            if (gray < 30)
                            {
                                totalDarkPixels++;
                                if (x < minX) minX = x;
                                if (x > maxX) maxX = x;
                            }
                        }
                    }

                    // Nếu tìm thấy đủ diện tích vùng tối (tương đương điều kiện area > 100 của bạn)
                    if (totalDarkPixels > 100 && minX <= maxX)
                    {
                        holeX = minX; // Lấy tọa độ X bắt đầu của vùng thủng/tối
                    }
                }

                return holeX;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi FindHolePosition: {ex.Message}");
                return -1;
            }
        }
        static int FindHolePosition(string imagePath)
        {
            try
            {
                Image<Bgr, byte> img = new Image<Bgr, byte>(imagePath);
                Image<Gray, byte> gray = img.Convert<Gray, byte>();

                Image<Gray, byte> darkMask = gray.ThresholdBinaryInv(new Gray(30), new Gray(255));

                VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
                Mat hierarchy = new Mat();
                CvInvoke.FindContours(darkMask, contours, hierarchy, RetrType.External, ChainApproxMethod.ChainApproxSimple);

                int holeX = -1;
                int maxArea = 0;

                for (int i = 0; i < contours.Size; i++)
                {
                    double area = CvInvoke.ContourArea(contours[i]);
                    if (area > maxArea && area > 100)
                    {
                        maxArea = (int)area;
                        System.Drawing.Rectangle rect = CvInvoke.BoundingRectangle(contours[i]);
                        holeX = rect.X;
                    }
                }

                img.Dispose();
                gray.Dispose();
                darkMask.Dispose();
                contours.Dispose();
                hierarchy.Dispose();

                return holeX;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi FindHolePosition: {ex.Message}");
                return -1;
            }
        }
        private async Task TraCuuNCC(bool toggle, string mst, string mabm, string shdon,string savedPath,int type)
        {
            string year = dateEdit1.DateTime.Year.ToString();
            string month = dateEdit1.DateTime.Month.ToString();

            await Task.Run(async () =>
            {
                string baseDir;
                if (type == 1)
                    baseDir = Path.Combine(savedPath, $"HD{year}", "HDVao", month);
                else if (type == 2)
                    baseDir = Path.Combine(savedPath, $"HD{year}", "HDRa", month);
                else
                    baseDir = Path.Combine(savedPath, $"HD{year}", "HDKhac", month);

                Directory.CreateDirectory(baseDir);

                // Thư mục tạm để Chrome tải file về (tách riêng khỏi thư mục lưu chính thức)
                string tempDownloadFolder = Path.Combine(Path.GetTempPath(), "MisaTemp", Guid.NewGuid().ToString("N"));

                // 2. Tạo thư mục thực tế trên ổ đĩa
                Directory.CreateDirectory(tempDownloadFolder);

                // [SỬA QUAN TRỌNG]: Chrome trên .NET Framework yêu cầu đổi dấu \ thành / để tránh lỗi "path is not of a legal form"
                string chromeDownloadPath = tempDownloadFolder.Replace("\\", "/");

                // Cấu hình ChromeOptions
                // Cấu hình ChromeOptions
                ChromeOptions options = new ChromeOptions();

                // === CHẠY ẨN (HEADLESS) & CHỐNG CHẶN TẢI FILE ===
                //  options.AddArgument("--headless");
                options.AddArgument("--disable-gpu");
                options.AddArgument("--no-sandbox");
                options.AddArgument("--disable-dev-shm-usage");
                options.AddArgument("--window-size=1920,1080");
                options.AddArgument("--disable-extensions");
                options.AddArgument("--disable-popup-blocking");
                options.AddArgument("--disable-web-security");

                // Cấu hình tải file
                options.AddUserProfilePreference("download.default_directory", tempDownloadFolder);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("safebrowsing.enabled", true);
                options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);

                IWebDriver driver = null;
                try
                {
                    string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                    service.HideCommandPromptWindow = true; // Ẩn cửa sổ cmd đen

                    driver = new ChromeDriver(service, options);
                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                    // ═══════════════════════════════════════════════════════
                    // 2. TRUY CẬP TRANG
                    // ═══════════════════════════════════════════════════════
                    driver.Navigate().GoToUrl("https://vinvoice.viettel.vn/utilities/invoice-search");
                    await Task.Delay(1500);

                     wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                    // ═══════════════════════════════════════════════════════
                    // 3. NHẬP MST + MÃ TRA CỨU
                    // ═══════════════════════════════════════════════════════
                    try
                    {
                        var taxInput = wait.Until(d =>
                        {
                            try
                            {
                                var el = d.FindElement(By.CssSelector("input[formcontrolname='supplierTaxCode']"));
                                return (el != null && el.Displayed) ? el : null;
                            }
                            catch { return null; }
                        });

                        taxInput.Clear();
                        taxInput.SendKeys(mst);

                        var codeInput = wait.Until(d =>
                        {
                            try
                            {
                                var el = d.FindElement(By.CssSelector("input[formcontrolname='reservationCode']"));
                                return (el != null && el.Displayed) ? el : null;
                            }
                            catch { return null; }
                        });

                        codeInput.Clear();
                        codeInput.SendKeys(mabm);

                        Console.WriteLine("✅ Đã nhập MST + mã tra cứu");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Không nhập được form: {ex.Message}");
                        return;
                    }

                    // ═══════════════════════════════════════════════════════
                    // 4. XỬ LÝ CAPTCHA — TỐI ĐA 3 LẦN
                    // ═══════════════════════════════════════════════════════
                    bool captchaSuccess = false;
                    int maxCaptchaRetries = 3;
                    int captchaAttempt = 0;

                    while (captchaAttempt < maxCaptchaRetries && !captchaSuccess)
                    {
                        captchaAttempt++;
                        Console.WriteLine($"🔵 Captcha lần {captchaAttempt}/{maxCaptchaRetries}");

                        try
                        {
                            // Đợi captcha load — kiểm tra loại captcha
                            await Task.Delay(500);

                            // ═══ CHẨN ĐOÁN: Kiểm tra loại captcha ═══
                            bool hasSlider = driver.FindElements(By.CssSelector("div.slider-button")).Count > 0;
                            bool hasRecaptcha = driver.FindElements(By.CssSelector("iframe[src*='recaptcha'], div.g-recaptcha")).Count > 0;
                            bool hasHcaptcha = driver.FindElements(By.CssSelector("iframe[src*='hcaptcha'], div.h-captcha")).Count > 0;
                            bool hasTurnstile = driver.FindElements(By.CssSelector("iframe[src*='cloudflare'], div.cf-turnstile")).Count > 0;

                            Console.WriteLine($"🔍 Captcha: slider={hasSlider}, recaptcha={hasRecaptcha}, hcaptcha={hasHcaptcha}, turnstile={hasTurnstile}");

                            // Không có captcha nào → tiếp tục
                            if (!hasSlider && !hasRecaptcha && !hasHcaptcha && !hasTurnstile)
                            {
                                Console.WriteLine("⚠️ Không phát hiện captcha — tiếp tục luôn");
                                captchaSuccess = true;
                                break;
                            }

                            // Captcha loại khác (không phải slider) → không tự xử lý được
                            if (hasRecaptcha || hasHcaptcha || hasTurnstile)
                            {
                                Console.WriteLine("⚠️ Captcha đã đổi loại (reCAPTCHA/hCaptcha/Turnstile) — cần user nhập thủ công");
                                // Không tự động được, nhưng vẫn thử tiếp để user có thời gian nhập
                                await Task.Delay(500);
                                captchaSuccess = true;
                                break;
                            }

                            // ═══ XỬ LÝ SLIDER CAPTCHA ═══
                            if (hasSlider)
                            {
                                // Nếu là lần thử lại → refresh và nhập lại
                                if (captchaAttempt > 1)
                                {
                                    Console.WriteLine("🔄 Refresh + nhập lại MST + mã...");
                                    driver.Navigate().Refresh();
                                    await Task.Delay(500);

                                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                                    try
                                    {
                                        var taxRetry = wait.Until(d =>
                                        {
                                            try
                                            {
                                                var el = d.FindElement(By.CssSelector("input[formcontrolname='supplierTaxCode']"));
                                                return el.Displayed ? el : null;
                                            }
                                            catch { return null; }
                                        });
                                        taxRetry.Clear();
                                        taxRetry.SendKeys(mst);

                                        var codeRetry = wait.Until(d =>
                                        {
                                            try
                                            {
                                                var el = d.FindElement(By.CssSelector("input[formcontrolname='reservationCode']"));
                                                return el.Displayed ? el : null;
                                            }
                                            catch { return null; }
                                        });
                                        codeRetry.Clear();
                                        codeRetry.SendKeys(mabm);
                                    }
                                    catch { }

                                    await Task.Delay(500);
                                }

                                // Tìm slider button
                                var bgImage = driver.FindElements(By.CssSelector("img.bg-image"));
                                if (bgImage.Count == 0)
                                {
                                    Console.WriteLine("⚠️ Không tìm thấy img.bg-image — captcha có thể chưa load xong");
                                    await Task.Delay(500);
                                    continue;
                                }

                                string bgSrc = bgImage[0].GetAttribute("src");
                                if (string.IsNullOrEmpty(bgSrc))
                                {
                                    Console.WriteLine("⚠️ Không lấy được src của bg-image");
                                    await Task.Delay(500);
                                    continue;
                                }

                                string bgPath = SaveBase64ToFile(bgSrc, "bg");
                                int holeX = FindHolePosition2(bgPath);
                                Console.WriteLine($"🎯 Ô đen ở vị trí: {holeX}px");

                                if (holeX <= 0)
                                {
                                    Console.WriteLine("⚠️ Không tìm được vị trí ô đen");
                                    try { File.Delete(bgPath); } catch { }
                                    await Task.Delay(500);
                                    continue;
                                }

                                // Tìm slider button + bar
                                var sliderButtons = driver.FindElements(By.CssSelector("div.slider-button"));
                                var sliderBars = driver.FindElements(By.CssSelector("div.slider-bar"));

                                if (sliderButtons.Count == 0 || sliderBars.Count == 0)
                                {
                                    Console.WriteLine("⚠️ Không tìm thấy slider-button hoặc slider-bar");
                                    try { File.Delete(bgPath); } catch { }
                                    await Task.Delay(500);
                                    continue;
                                }

                                var sliderButton = sliderButtons[0];
                                var sliderBar = sliderBars[0];

                                int barStartX = sliderBar.Location.X;
                                int buttonStartX = sliderButton.Location.X;
                                int offset = buttonStartX - barStartX;
                                int targetDistance = holeX - offset;

                                // Kéo slider
                                Actions actions = new Actions(driver);
                                actions.ClickAndHold(sliderButton).Perform();

                                int steps = Math.Max(2, targetDistance / 30);
                                int stepDistance = targetDistance / steps;
                                int remaining = targetDistance % steps;

                                for (int i = 0; i < steps; i++)
                                {
                                    actions.MoveByOffset(stepDistance, 0).Perform();
                                    await Task.Delay(10);
                                }

                                if (remaining > 0)
                                {
                                    actions.MoveByOffset(remaining, 0).Perform();
                                }

                                actions.Release().Perform();

                                // Đợi DOM cập nhật
                                await Task.Delay(1000);

                                // Check thành công đa tiêu chí
                                captchaSuccess = CheckCaptchaSuccess(driver);

                                try { File.Delete(bgPath); } catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Lỗi captcha lần {captchaAttempt}: {ex.Message}");
                        }

                        if (!captchaSuccess)
                        {
                            await Task.Delay(1000);
                        }
                    }

                    if (!captchaSuccess)
                    {
                        Console.WriteLine("❌ Thử captcha 3 lần không thành công.");
                        return;
                    }

                    Console.WriteLine("✅ Captcha OK — chuẩn bị tìm kiếm");
                    await Task.Delay(800);

                    // ═══════════════════════════════════════════════════════
                    // 5. CLICK NÚT TÌM KIẾM
                    // ═══════════════════════════════════════════════════════
                    bool searchSuccess = false;
                    try
                    {
                        var searchBtns = driver.FindElements(By.XPath(
                            "//button[@type='submit' and (contains(., 'Tìm kiếm') or .//span[contains(text(), 'Tìm kiếm')])]"));

                        if (searchBtns.Count == 0)
                        {
                            Console.WriteLine("❌ Không tìm thấy nút Tìm kiếm");
                            return;
                        }

                        var searchBtn = searchBtns[0];
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center'});", searchBtn);
                        await Task.Delay(300);
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", searchBtn);

                        Console.WriteLine("✅ Đã click Tìm kiếm");
                        await Task.Delay(2000);

                        // Đợi nút tải PDF xuất hiện = có kết quả
                        var pdfWaitCheck = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                        try
                        {
                            pdfWaitCheck.Until(d =>
                            {
                                var btns = d.FindElements(By.XPath(
                                    "//button[.//span[contains(text(),'Tải về file PDF')]]"));
                                return btns.Count > 0 ? btns[0] : null;
                            });
                            searchSuccess = true;
                            Console.WriteLine("✅ Có kết quả tìm kiếm");
                        }
                        catch
                        {
                            Console.WriteLine("⚠️ Không thấy kết quả — có thể không có hóa đơn hoặc captcha sai");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Lỗi tìm kiếm: {ex.Message}");
                        return;
                    }

                    if (!searchSuccess)
                    {
                        Console.WriteLine("❌ Không có kết quả để tải");
                        return;
                    }

                    // ═══════════════════════════════════════════════════════
                    // 6. TẢI PDF
                    // ═══════════════════════════════════════════════════════
                    try
                    {
                        var strategies = new List<Func<IWebDriver, IWebElement>>
                {
                    d => d.FindElement(By.XPath("//button[.//span[contains(text(),'Tải về file PDF')]]")),
                    d => d.FindElement(By.XPath("//button[.//i[contains(@class, 'fa-file-pdf-o')]]")),
                    d => d.FindElement(By.XPath("//button[contains(@class, 'btn-link') and contains(., 'PDF')]")),
                    d => d.FindElement(By.XPath("//a[contains(text(),'Tải về file PDF')]")),
                };

                        var pdfWait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                        IWebElement pdfBtn = pdfWait.Until(d =>
                        {
                            foreach (var strategy in strategies)
                            {
                                try
                                {
                                    var el = strategy(d);
                                    if (el != null && el.Displayed && el.Enabled) return el;
                                }
                                catch { }
                            }
                            return null;
                        });

                        if (pdfBtn == null)
                        {
                            Console.WriteLine("❌ Không tìm thấy nút tải PDF");
                            return;
                        }

                        Console.WriteLine("📥 Đang click tải PDF...");

                        var beforeFiles = new HashSet<string>(
                            Directory.GetFiles(tempDownloadFolder, "*.*").Select(Path.GetFileName)
                        );

                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center'});", pdfBtn);
                        await Task.Delay(300);
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", pdfBtn);

                        Console.WriteLine("✅ Đã click tải PDF, chờ file...");

                        // Chờ file PDF xuất hiện
                        string latestPdfFile = null;
                        var deadline = DateTime.Now.AddSeconds(20);

                        while (DateTime.Now < deadline)
                        {
                            var newFiles = Directory.GetFiles(tempDownloadFolder, "*.pdf")
                                .Where(f => !beforeFiles.Contains(Path.GetFileName(f)))
                                .ToList();

                            if (newFiles.Count > 0)
                            {
                                var candidate = newFiles.OrderByDescending(File.GetCreationTime).First();

                                try
                                {
                                    using (var fs = File.Open(candidate, FileMode.Open, FileAccess.Read, FileShare.None)) { }
                                    latestPdfFile = candidate;
                                    break;
                                }
                                catch (IOException) { }
                            }

                            await Task.Delay(200);
                        }

                        if (latestPdfFile == null)
                        {
                            Console.WriteLine("❌ Không tải được PDF hoặc quá thời gian");
                            return;
                        }

                        // Đổi tên file
                        string safeMst = MakeSafeFileName(mst);
                        string safeShdon = MakeSafeFileName(shdon);
                        string newPath = Path.Combine(baseDir, $"{safeMst}_{safeShdon}.pdf");

                        int counter = 1;
                        string dir = Path.GetDirectoryName(newPath);
                        string nameNoExt = Path.GetFileNameWithoutExtension(newPath);
                        string ext = Path.GetExtension(newPath);

                        while (File.Exists(newPath))
                        {
                            newPath = Path.Combine(dir, $"{nameNoExt}_{counter}{ext}");
                            counter++;
                        }

                        File.Move(latestPdfFile, newPath);
                        Console.WriteLine($"📁 Đã lưu: {newPath}");

                        if (toggle)
                        {
                            try
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                {
                                    FileName = newPath,
                                    UseShellExecute = true
                                });
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Lỗi tải PDF: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Lỗi tổng: {ex.Message}");
                }
                finally
                {
                    try { driver?.Quit(); } catch { }
                    try { driver?.Dispose(); } catch { }

                    try
                    {
                        if (Directory.Exists(tempDownloadFolder))
                            Directory.Delete(tempDownloadFolder, true);
                    }
                    catch { }
                }
            });
        }
        private bool CheckCaptchaSuccess(IWebDriver driver)
        {
            try
            {
                // Check 1: Button biến mất
                var sliderBtns = driver.FindElements(By.CssSelector("div.slider-button"));
                if (sliderBtns.Count == 0)
                {
                    Console.WriteLine("   ✅ Check 1: Button biến mất");
                    return true;
                }

                if (!sliderBtns[0].Displayed)
                {
                    Console.WriteLine("   ✅ Check 1b: Button không hiển thị");
                    return true;
                }

                // Check 2: Class success
                string btnClass = sliderBtns[0].GetAttribute("class") ?? "";
                string btnText = sliderBtns[0].Text?.Trim() ?? "";

                if (btnClass.IndexOf("success", StringComparison.OrdinalIgnoreCase) >= 0
                    || btnClass.IndexOf("verified", StringComparison.OrdinalIgnoreCase) >= 0
                    || btnClass.IndexOf("complete", StringComparison.OrdinalIgnoreCase) >= 0
                    || btnClass.IndexOf("is-success", StringComparison.OrdinalIgnoreCase) >= 0
                    || btnText.Contains("thành công")
                    || btnText.Contains("xác nhận")
                    || btnText.Contains("✓"))
                {
                    Console.WriteLine($"   ✅ Check 2: class='{btnClass}', text='{btnText}'");
                    return true;
                }

                // Check 3: Progress bar 100%
                try
                {
                    var progresses = driver.FindElements(By.CssSelector(
                        "div.slider-progress, div.slider-move, div.slider-track, div.slider-indicator"));

                    foreach (var p in progresses)
                    {
                        try
                        {
                            if (!p.Displayed) continue;
                            string style = p.GetAttribute("style") ?? "";

                            if (style.IndexOf("100%", StringComparison.OrdinalIgnoreCase) >= 0
                                || style.IndexOf("scaleX(1)", StringComparison.OrdinalIgnoreCase) >= 0
                                || style.IndexOf("translateX(100%", StringComparison.OrdinalIgnoreCase) >= 0
                                || style.IndexOf("translate3d(100%", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                Console.WriteLine($"   ✅ Check 3: progress='{style}'");
                                return true;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // Check 4: Text success
                try
                {
                    var texts = driver.FindElements(By.XPath(
                        "//*[contains(text(),'Xác nhận thành công') or contains(text(),'xác thực thành công') or contains(text(),'Verified')]"));

                    foreach (var el in texts)
                    {
                        try
                        {
                            if (el.Displayed && !string.IsNullOrEmpty(el.Text))
                            {
                                Console.WriteLine($"   ✅ Check 4: text='{el.Text}'");
                                return true;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // Check 5: Nút Tìm kiếm enabled
                try
                {
                    var searchBtn = driver.FindElements(By.XPath(
                        "//button[@type='submit' and (contains(., 'Tìm kiếm') or .//span[contains(text(), 'Tìm kiếm')])]"));

                    if (searchBtn.Count > 0 && searchBtn[0].Enabled)
                    {
                        Console.WriteLine("   ✅ Check 5: Nút Tìm kiếm enabled");
                        return true;
                    }
                }
                catch { }

                Console.WriteLine("   ⏭ Chưa phát hiện thành công");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ Lỗi CheckCaptchaSuccess: {ex.Message}");
                return false;
            }
        }
        private static string MakeSafeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "unknown";

            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            return name.Trim();
        }
        public static string GetSecretCode(string xmlFilePath, string nhaCungCap)
        {
            if (!File.Exists(xmlFilePath))
                return "❌ File không tồn tại!";

            try
            {
                XDocument doc = XDocument.Load(xmlFilePath);
                XNamespace ns = doc.Root.GetDefaultNamespace();

                // Chuẩn hóa tên nhà cung cấp để so sánh
                string provider = nhaCungCap.ToLower().Trim();
                string maTraCuu = null;

                switch (provider)
                {
                    case "viettel":
                        maTraCuu = FindViettelCode(doc, ns);
                        break;
                    case "misa":
                    case "meinvoice":
                        maTraCuu = FindMisaCode(doc, ns);
                        break;
                    case "bkav":
                        maTraCuu = FindBkavCode(doc, ns);
                        break;
                    case "vnpt":
                        maTraCuu = FindVnptCode(doc, ns);
                        break;
                    case "minvoice":
                        maTraCuu = FindMInvoiceCode(doc, ns);
                        break;
                    case "win tech":
                    case "winvoice":
                        maTraCuu = FindWinTechCode(doc, ns);
                        break;
                    case "fast":
                        maTraCuu = FindFastCode(doc, ns);
                        break;
                    case "cyber":
                        maTraCuu = FindCyberCode(doc, ns);
                        break;
                    case "softdream":
                        maTraCuu = FindSoftdream(doc, ns);
                        break;
                    case "visnam":
                        maTraCuu = FindVisnamCode(doc, ns);
                        break;
                    default:
                        // Nếu không xác định được NCC, thử tìm tất cả các loại mã
                        maTraCuu = FindAllPossibleCodes(doc, ns);
                        break;
                }

                return string.IsNullOrEmpty(maTraCuu) ? "❌ Không tìm thấy mã tra cứu!" : maTraCuu;
            }
            catch (Exception ex)
            {
                return $"❌ Lỗi đọc XML: {ex.Message}";
            }
        }
        private static string FindViettelCode(XDocument doc, XNamespace ns)
        {
            // Ưu tiên tìm "Mã số bí mật" trước
            string maBiMat = FindElementByTagName(doc, "Mã số bí mật");
            if (!string.IsNullOrEmpty(maBiMat)) return maBiMat;

            // Nếu không có, thử các trường khác
            return FindElementValue(doc, ns, "KeySearch") ??
                   FindElementValue(doc, ns, "reservationCode");
        }
        private static string FindMisaCode(XDocument doc, XNamespace ns)
        {
            var dLHDON = doc.Descendants("DLHDon").FirstOrDefault();
            if (dLHDON != null)
            {
                XAttribute idAttr = dLHDON.Attribute("Id");
                if (idAttr != null && !string.IsNullOrEmpty(idAttr.Value))
                {
                    return idAttr.Value; // Trả về: W1F9TZ_9ZNM0
                }
            }
            // MISA meInvoice: KeySearch hoặc TransactionID
            return FindElementValue(doc, ns, "KeySearch") ??
                   FindElementValue(doc, ns, "TransactionID");
        }

        private static string FindBkavCode(XDocument doc, XNamespace ns)
        {

            // DLHDon thường là con trực tiếp của root <HDon>
            var dlhdon = doc.Root?
                .Elements()
                .FirstOrDefault(e => e.Name.LocalName == "DLHDon");

            var idAttr = dlhdon?
                .Attributes()
                .FirstOrDefault(a => a.Name.LocalName == "Id");

            if (idAttr != null && !string.IsNullOrWhiteSpace(idAttr.Value))
                return idAttr.Value.Trim();

            return null;
        }

        private static string FindVnptCode(XDocument doc, XNamespace ns)
        {
            // VNPT: ưu tiên MCCQT
            var mccqt = FindElementValue(doc, ns, "MCCQT");
            if (!string.IsNullOrWhiteSpace(mccqt))
                return mccqt.Trim();

            var keySearch = FindElementValue(doc, ns, "KeySearch");
            if (!string.IsNullOrWhiteSpace(keySearch))
                return keySearch.Trim();

            var secretCode = FindElementValue(doc, ns, "SecretCode");
            if (!string.IsNullOrWhiteSpace(secretCode))
                return secretCode.Trim();

            var maBaoMat = FindElementByTagName(doc, "Mã bảo mật");
            return string.IsNullOrWhiteSpace(maBaoMat) ? null : maBaoMat.Trim();
        }

        private static string FindMInvoiceCode(XDocument doc, XNamespace ns)
        {
            // M-Invoice: Số bảo mật hoặc Mã tra cứu
            return FindElementByTagName(doc, "Số bảo mật") ??
                   FindElementByTagName(doc, "Mã tra cứu") ??
                   FindElementValue(doc, ns, "TransactionID");
        }

        private static string FindWinTechCode(XDocument doc, XNamespace ns)
        {
            // WinTech: Mã tra cứu hoặc privateCode
            return FindElementByTagName(doc, "Mã tra cứu hóa đơn") ??    // Tìm theo attribute Id
                   FindElementValue(doc, ns, "TransactionID");
        }
        private static string FindElementValue(XDocument doc, XNamespace ns, string elementName)
        {
            // Tìm theo namespace
            var element = doc.Descendants(ns + elementName).FirstOrDefault();
            if (element != null)
                return element.Value;

            // Tìm không namespace
            var element2 = doc.Descendants(elementName).FirstOrDefault();
            return element2?.Value;
        }

        private static string FindElementByTagName(XDocument doc, string tagName)
        {
            // Tìm cấu trúc TTruong + DLieu tương ứng
            var tTinElements = doc.Descendants("TTin");
            foreach (var tTin in tTinElements)
            {
                var truong = tTin.Descendants("TTruong").FirstOrDefault();
                if (truong != null && truong.Value == tagName)
                {
                    var dLieu = tTin.Descendants("DLieu").FirstOrDefault();
                    if (dLieu != null && !string.IsNullOrEmpty(dLieu.Value))
                        return dLieu.Value;
                }
            }
            return null;
        }
        private static string FindCyberCode(XDocument doc, XNamespace ns)
        {
            var maTraCuu = doc.Descendants("TTin")
                .FirstOrDefault(x => (string)x.Element("TTruong") == "MaTraCuu")
                ?.Element("DLieu")?.Value;

            return maTraCuu?.Trim();
        }
        //FindVisnamCode
        private static string FindVisnamCode(XDocument doc, XNamespace ns)
        {
            var maTraCuu = doc.Descendants("TTin")
                .FirstOrDefault(x => (string)x.Element("TTruong") == "QuanLy_SoBaoMat")
                ?.Element("DLieu")?.Value;

            return maTraCuu?.Trim();
        }
        private static string FindSoftdream(XDocument doc, XNamespace ns)
        {
            var maTraCuu = doc.Descendants("TTin")
                .FirstOrDefault(x => (string)x.Element("TTruong") == "Fkey")
                ?.Element("DLieu")?.Value;

            return maTraCuu?.Trim();
        }
        private static string FindFastCode(XDocument doc, XNamespace ns)
        {
            // Tìm tất cả <TTin> ở mọi cấp, lọc theo <TTruong> = "KeySearch"
            var ttin = doc.Descendants()
                .FirstOrDefault(e =>
                    e.Name.LocalName == "TTin" &&
                    e.Elements().Any(c =>
                        c.Name.LocalName == "TTruong" &&
                        c.Value.Trim() == "KeySearch"));

            if (ttin == null) return null;

            var dlieu = ttin.Elements()
                .FirstOrDefault(c => c.Name.LocalName == "DLieu");

            return dlieu?.Value?.Trim();
        }

        private static string FindAllPossibleCodes(XDocument doc, XNamespace ns)
        {
            // Thử tìm tất cả các loại mã thường gặp
            string[] possibleTags = new string[]
            {
            "KeySearch", "TransactionID", "reservationCode", "SecretCode",
            "Mã tra cứu", "Số bảo mật", "Mã bảo mật", "privateCode"
            };

            foreach (var tag in possibleTags)
            {
                string value = FindElementByTagName(doc, tag) ?? FindElementValue(doc, ns, tag);
                if (!string.IsNullOrEmpty(value))
                    return value;
            }
            return null;
        }
        public async Task GetKNMXMLAsync(string nbmst, string khhdon, string shdon, string token,
            DateTime GetNLap, string path, string filename, string companyName)
        {
            string url = $"{AppConstants.GdtBaseUrl}/api/query/invoices/detail?nbmst={nbmst}&khhdon={khhdon}&shdon={shdon}&khmshdon=1";

            try
            {
                var client = GdtHttpClientFactory.Instance;
                 var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                 var response = await client.SendAsync(req);

                if (response.StatusCode == HttpStatusCode.Unauthorized)
                {
                    Log($"❌ {companyName}: Token hết hạn khi lấy KNM XML cho HĐ {shdon}");
                    return;
                }

                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    if (string.IsNullOrEmpty(responseBody))
                    {
                        Log($"⚠️ {companyName}: Response rỗng cho HĐ {shdon}");
                        return;
                    }

                    var rootObject = JsonConvert.DeserializeObject<Invoice>(responseBody);
                    TaoFileXmlChiCoDLHDon(path, filename.Replace(".zip", ""), rootObject, GetNLap);
                    Log($"✅ {companyName}: Đã tạo KNM XML cho HĐ {shdon}");
                }
            }
            catch (Exception ex)
            {
                Log($"❌ {companyName}: Lỗi GetKNMXMLAsync: {ex.Message}");
            }
        }

        #endregion

        #region Utility Methods

        public static string RemoveLeadingZeros(string invoiceNumber)
        {
            if (string.IsNullOrEmpty(invoiceNumber)) return invoiceNumber;
            return RegexPatterns.LeadingZeros.Replace(invoiceNumber, "");
        }

        private static void ExtractZipXML(string path)
        {
            try
            {
                string rootPath = Path.GetDirectoryName(path);
                string getnamefile = Path.GetFileNameWithoutExtension(path);
                string directoryPath = rootPath + @"\Giainen_" + getnamefile;

                ZipFile.ExtractToDirectory(path, directoryPath);

                var files = Directory.GetFiles(directoryPath, "invoice.html", SearchOption.AllDirectories);
                if (files.Any())
                    File.Move(files.FirstOrDefault(), Path.Combine(rootPath, getnamefile + ".html"));

                var filesxml = Directory.GetFiles(directoryPath, "invoice.xml", SearchOption.AllDirectories);
                if (filesxml.Any())
                    File.Move(filesxml.FirstOrDefault(), Path.Combine(rootPath, getnamefile + ".xml"));

                File.Delete(path);
                Directory.Delete(directoryPath, true);
            }
            catch (Exception ex) { Log($"Lỗi giải nén: {ex.Message}"); }
        }

        public static double SafeParse(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return 0;
            value = value.Replace(",", ".");
            if (double.TryParse(value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double result))
                return result;
            return 0;
        }

        public static string NormalizeVietnameseString(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            input = input.Replace("*", "x").Replace(" lít", "l").Replace("( ", "(").Replace("\t", "");
            return input.Normalize(NormalizationForm.FormC).Trim();
        }

        public static string GenerateResultString(string input)
        {
            string firstWord = input.Split(' ')[0];
            string normalizedFirstWord = RemoveVietnameseDiacritics(firstWord).Replace("á", "a");
            if (normalizedFirstWord.Length >= 10)
            {
                if (normalizedFirstWord.Length > 10) normalizedFirstWord = normalizedFirstWord.Substring(0, 10);
                normalizedFirstWord = char.ToUpper(normalizedFirstWord[0]) + normalizedFirstWord.Substring(1);
            }
            string randomNumbers = GenerateRandomNumbers(4);
            return CapitalizeFirstLetter(normalizedFirstWord).ToUpper() + randomNumbers;
        }

        private static string RemoveVietnameseDiacritics(string str)
        {
            str = str.ToLower();
            str = Regex.Replace(str, "[àáạảãâầấậẩẫăằắặẳẵ]", "a");
            str = Regex.Replace(str, "[èéẹẻẽêềếệểễ]", "e");
            str = Regex.Replace(str, "[ìíịỉĩ]", "i");
            str = Regex.Replace(str, "[òóọỏõôồốộổỗơờớợởỡ]", "o");
            str = Regex.Replace(str, "[ùúụủũưừứựửữ]", "u");
            str = Regex.Replace(str, "[ỳýỵỷỹ]", "y");
            str = Regex.Replace(str, "đ", "d");
            str = str.Replace(",", "").Replace(".", "").Replace("*", "x").Replace("'", "");
            return str;
        }

        static string CapitalizeFirstLetter(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            return char.ToUpper(input[0]) + input.Substring(1);
        }

        private static Random random = new Random();

        private static string GenerateRandomNumbers(int length)
        {
            string randomNumbers = "";
            HashSet<int> generated = new HashSet<int>();
            while (randomNumbers.Length < length)
            {
                int number = random.Next(1, 10);
                if (!generated.Contains(number))
                {
                    randomNumbers += number.ToString();
                    generated.Add(number);
                }
            }
            return randomNumbers;
        }

        public string GenerateAbbreviation(string fullName, List<string> existingNames)
        {
            string[] nameParts = fullName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string abbreviation = "";
            foreach (string part in nameParts) abbreviation += part[0].ToString().ToLower();

            int counter = 1;
            string uniqueAbbreviation = abbreviation.ToUpper();
            while (existingNames.Contains(uniqueAbbreviation))
            {
                uniqueAbbreviation = abbreviation + "-" + counter;
                counter++;
            }
            return uniqueAbbreviation;
        }

        private static void ExtractZipXMLAsynce(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    XtraMessageBox.Show("File không tồn tại: " + path);
                    return;
                }
                ExtractZipXML(path);
            }
            catch (Exception ex) { Console.WriteLine($"Lỗi giải nén: {ex.Message}"); }
        }

        #endregion

        #region XML File Creation

        public static string TaoFileXmlChiCoDLHDon(string folderPath, string fileName = "HDon_44020.xml",
            Invoice Invoice = default, DateTime NLap = default)
        {
            if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);
            fileName = $"{fileName}_KNM.xml";
            string fullPath = Path.Combine(folderPath, fileName);

            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("HDon");
            doc.AppendChild(root);

            XmlElement dlHDon = doc.CreateElement("DLHDon");
            dlHDon.SetAttribute("Id", "DuLieuKy");
            root.AppendChild(dlHDon);

            XmlElement ttChung = doc.CreateElement("TTChung");
            dlHDon.AppendChild(ttChung);
            ThemPhanTu(doc, ttChung, "PBan", $"{Invoice.Pban}");
            ThemPhanTu(doc, ttChung, "THDon", "Hóa đơn GTGT");
            ThemPhanTu(doc, ttChung, "KHMSHDon", $"{Invoice.Khmshdon}");
            ThemPhanTu(doc, ttChung, "KHHDon", $"{Invoice.Khhdon}");
            ThemPhanTu(doc, ttChung, "SHDon", $"{Invoice.Shdon}");
            ThemPhanTu(doc, ttChung, "NLap", $"{NLap}");
            ThemPhanTu(doc, ttChung, "HDCTTChinh", "0");
            ThemPhanTu(doc, ttChung, "DVTTe", $"{Invoice.Dvtte}");
            ThemPhanTu(doc, ttChung, "TGia", $"{Invoice.Tgia}");
            ThemPhanTu(doc, ttChung, "HTTToan", $"{Invoice.Thtttoan}");
            ThemPhanTu(doc, ttChung, "MSTTCGP", $"{Invoice.Msttcgp}");

            XmlElement ttKhacChung = doc.CreateElement("TTKhac");
            ttKhacChung.AppendChild(TaoTTin(doc, "Extra", "string", "ThuyUyen.Nguyen@britishcouncil.org"));
            ttChung.AppendChild(ttKhacChung);

            XmlElement ndHDon = doc.CreateElement("NDHDon");
            dlHDon.AppendChild(ndHDon);

            XmlElement nBan = doc.CreateElement("NBan");
            ThemPhanTu(doc, nBan, "Ten", $"{Invoice.Nbten}");
            ThemPhanTu(doc, nBan, "MST", $"{Invoice.Nbmst}");
            ThemPhanTu(doc, nBan, "DChi", $"{Invoice.Nbdchi}");
            ThemPhanTu(doc, nBan, "SDThoai", $"{Invoice.Nbsdthoai}");
            ndHDon.AppendChild(nBan);

            XmlElement nMua = doc.CreateElement("NMua");
            ThemPhanTu(doc, nMua, "Ten", $"{Invoice.Nmten}");
            ThemPhanTu(doc, nMua, "MST", $"{Invoice.Nmmst}");
            ThemPhanTu(doc, nMua, "DChi", $"{Invoice.Nmdchi}");
            ThemPhanTu(doc, nMua, "MKHang", $"{Invoice.Mkhang}");
            ThemPhanTu(doc, nMua, "HVTNMHang", "");
            ndHDon.AppendChild(nMua);

            XmlElement dsHHDVu = doc.CreateElement("DSHHDVu");
            ndHDon.AppendChild(dsHHDVu);

            int stt = 1;
            foreach (var dt in Invoice.Hdhhdvu.ToList())
            {
                TaoHangHoa(doc, dsHHDVu, "0", $"{stt}",
                    !string.IsNullOrEmpty(dt.Ten) ? $"{dt.Ten}" : "Hoá đơn không nhận mã",
                    $"{dt.Dvtinh}", $"{dt.Sluong}", $"{dt.Dgia}", $"{dt.Tsuat.Value * 100}", $"{dt.Thtien}",
                    new[] { ("Amount", "numeric", $"{dt.Thtien}"), ("VATAmount", "numeric", "0") });
                stt++;
            }

            XmlElement tToan = doc.CreateElement("TToan");
            ndHDon.AppendChild(tToan);

            XmlElement ttKhacToan = doc.CreateElement("TTKhac");
            ttKhacToan.AppendChild(TaoTTin(doc, "ServiceProvided", "String", "Le phi thi"));
            ttKhacToan.AppendChild(TaoTTin(doc, "Location", "String", "British Council Ho Chi Minh City"));
            ttKhacToan.AppendChild(TaoTTin(doc, "Datasource", "String", "ORS2"));
            tToan.AppendChild(ttKhacToan);

            XmlElement tHTTLTSuat = doc.CreateElement("THTTLTSuat");
            XmlElement lTSuat = doc.CreateElement("LTSuat");
            ThemPhanTu(doc, lTSuat, "TSuat", $"{Invoice.Hdhhdvu.FirstOrDefault()?.Tsuat.Value * 100}");
            ThemPhanTu(doc, lTSuat, "TThue", $"{Invoice.Tgtthue}");
            ThemPhanTu(doc, lTSuat, "ThTien", $"{Invoice.Tgtcthue}");
            tHTTLTSuat.AppendChild(lTSuat);
            tToan.AppendChild(tHTTLTSuat);

            if (Invoice.Tgtphi.HasValue)
            {
                XmlElement TPhi = doc.CreateElement("TPhi");
                ThemPhanTu(doc, TPhi, "ThTien", $"{Invoice.Tgtphi}");
                tToan.AppendChild(TPhi);
            }

            ThemPhanTu(doc, tToan, "TgTCThue", $"{Invoice.Tgtcthue}");
            ThemPhanTu(doc, tToan, "TgTThue", $"{Invoice.Tgtthue}");
            ThemPhanTu(doc, tToan, "TTCKTMai", $"{Invoice.Ttcktmai}");
            ThemPhanTu(doc, tToan, "TgTTTBSo", $"{Invoice.Tgtttbso}");
            ThemPhanTu(doc, tToan, "TgTTTBChu", $"{Invoice.Tgtttbchu}");

            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                NewLineChars = "\r\n",
                Encoding = Encoding.UTF8
            };

            using (XmlWriter writer = XmlWriter.Create(fullPath, settings))
                doc.Save(writer);

            return fullPath;
        }

        static void ThemPhanTu(XmlDocument doc, XmlElement parent, string ten, string giaTri)
        {
            XmlElement e = doc.CreateElement(ten);
            e.InnerText = giaTri;
            parent.AppendChild(e);
        }

        static XmlElement TaoTTin(XmlDocument doc, string tTruong, string kDLieu, string dLieu)
        {
            XmlElement ttin = doc.CreateElement("TTin");
            ThemPhanTu(doc, ttin, "TTruong", tTruong);
            ThemPhanTu(doc, ttin, "KDLieu", kDLieu);
            ThemPhanTu(doc, ttin, "DLieu", dLieu);
            return ttin;
        }

        static void TaoHangHoa(XmlDocument doc, XmlElement ds, string tChat, string stt, string tenHang,
            string dvTinh, string sl, string dGia, string tSuat, string thTien,
            (string truong, string kieu, string giaTri)[] extra)
        {
            XmlElement hh = doc.CreateElement("HHDVu");
            ThemPhanTu(doc, hh, "TChat", tChat);
            ThemPhanTu(doc, hh, "STT", stt);
            ThemPhanTu(doc, hh, "THHDVu", tenHang);
            if (!string.IsNullOrEmpty(dvTinh)) ThemPhanTu(doc, hh, "DVTinh", dvTinh);
            ThemPhanTu(doc, hh, "SLuong", sl);
            ThemPhanTu(doc, hh, "DGia", dGia);
            ThemPhanTu(doc, hh, "TSuat", tSuat);
            ThemPhanTu(doc, hh, "ThTien", thTien);

            XmlElement ttKhac = doc.CreateElement("TTKhac");
            foreach (var item in extra)
                ttKhac.AppendChild(TaoTTin(doc, item.truong, item.kieu, item.giaTri));

            hh.AppendChild(ttKhac);
            ds.AppendChild(hh);
        }

        #endregion

        #region Logging

        private static void Log(string message)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} - {message}");
            if (_logAction != null) _logAction(message);
        }

        #endregion

        #region VatTu Loading

        public async Task<List<VatTu>> LoadDataVattuAsync(string conectionst)
        {
            List<VatTu> lstVattu = new List<VatTu>();

            try
            {
                var ListVattu = await Task.Run(() => ExecuteQuery2("SELECT * FROM Vattu", conectionst, null));
                ListPhanloaiVattu = await Task.Run(() => ExecuteQuery2("SELECT * FROM PhanLoaiVattu", conectionst, null));

                foreach (DataRow item in ListVattu.Rows)
                {
                    item["TenVattu"] = Helpers.ConvertVniToUnicode(item["TenVattu"].ToString());
                    item["TenVattu2"] = Helpers.ConvertVniToUnicode(item["TenVattu2"].ToString());
                    item["DonVi"] = Helpers.ConvertVniToUnicode(item["DonVi"].ToString());
                }

                var maVatTuList = ListVattu.Rows.Cast<DataRow>()
                    .Select(row => int.Parse(row["MaSo"].ToString())).Distinct().ToList();
                if (maVatTuList.Count == 0) return new List<VatTu>();

                var queryTonKhoBatch = @"SELECT * FROM TonKho WHERE MaVatTu IN (" + string.Join(",", maVatTuList) + ")";
                var allTonKho = await Task.Run(() => ExecuteQuery2(queryTonKhoBatch, conectionst, null));

                var tonKhoDict = allTonKho.Rows.Cast<DataRow>()
                    .GroupBy(row => int.Parse(row["MaVatTu"].ToString()))
                    .ToDictionary(group => group.Key, group => group.First());

                foreach (DataRow item in ListVattu.Rows)
                {
                    try
                    {
                        var maSo = int.Parse(item["MaSo"].ToString());
                        var tenMaPhanLoai = ListPhanloaiVattu.AsEnumerable()
                            .Where(m => m["MaSo"].ToString() == item["MaPhanLoai"].ToString())
                            .FirstOrDefault()?["TenPhanLoai"].ToString() ?? string.Empty;

                        var VatTu = new VatTu
                        {
                            MaSo = maSo,
                            MaPhanLoai = int.Parse(item["MaPhanLoai"].ToString()),
                            TenVattu = item["TenVattu"].ToString(),
                            TenVattu2 = item["TenVattu2"].ToString(),
                            SoHieu = item["SoHieu"].ToString(),
                            DonVi = item["DonVi"].ToString(),
                            GhiChu = item["GhiChu"].ToString(),
                            TenMaPhanLoai = tenMaPhanLoai,
                            PTGB = item["PTGB"].ToString(),
                        };

                        if (tonKhoDict.TryGetValue(maSo, out DataRow tonKhoRow))
                        {
                            int cnt = 12;
                            VatTu.SoLuong = tonKhoRow["Luong_" + cnt] != DBNull.Value
                                ? double.Parse(tonKhoRow["Luong_" + cnt].ToString()) : 0;
                            VatTu.ThanhTien = tonKhoRow["Tien_" + cnt] != DBNull.Value
                                ? double.Parse(tonKhoRow["Tien_" + cnt].ToString()) : 0;

                            if (VatTu.SoLuong != 0 && VatTu.ThanhTien != 0)
                                VatTu.Dongia = VatTu.ThanhTien / VatTu.SoLuong;

                            try
                            {
                                if (existingTbChungtu != null)
                                {
                                    var findlstct = existingTbChungtu.AsEnumerable()
                                        .Where(m => int.Parse(m["MaVattu"].ToString()) == VatTu.MaSo
                                                 && double.Parse(m["SoPS"].ToString()) != 0
                                                 && double.Parse(m["SoPS2Co"].ToString()) != 0
                                                 && !m["SoHieu"].ToString().Contains("V"))
                                        .LastOrDefault();
                                    if (findlstct != null)
                                    {
                                        double SoPS2Co = double.Parse(findlstct["SoPS2Co"].ToString());
                                        double SoPS = double.Parse(findlstct["SoPS"].ToString());
                                        VatTu.Dongia2 = SoPS2Co > 0 ? Math.Round(SoPS / SoPS2Co) : 0;
                                    }
                                }
                            }
                            catch (Exception ex) { XtraMessageBox.Show($"Lỗi tính đơn giá: {ex.Message}"); }
                        }
                        lstVattu.Add(VatTu);
                    }
                    catch (Exception ex) { XtraMessageBox.Show($"Lỗi vật tư: {ex.Message}"); }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi tải dữ liệu: {ex.Message}");
                throw;
            }

            BuildVatTuLookup(lstVattu);
            return lstVattu;
        }

        private void BuildVatTuLookup(List<VatTu> lstVattu)
        {
            _lookupByTenChinh = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _lookupByTenPhu = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _lookupByTenChinhs = new Dictionary<string, VatTuInfo>(StringComparer.OrdinalIgnoreCase);

            vatTuLookup = lstVattu.ToDictionary(v => v.SoHieu,
                v => (v.TenVattu, v.TenVattu2, v.DonVi, v.Dongia, v.SoLuong));

            var matcher = new VietnameseProductMatcher();
            foreach (var kvp in vatTuLookup)
            {
                string key1 = Helpers.NormalizeVietnameseString(kvp.Value.TenVattu.Trim());
                key1 = matcher.NormalizeVietnameseProduct(key1);

                if (!string.IsNullOrEmpty(key1))
                {
                    _lookupByTenChinh[key1] = kvp.Key;
                    _lookupByTenChinhs[key1] = new VatTuInfo { Ma = kvp.Key, DonViTinh = kvp.Value.DonVi };
                }

                if (!string.IsNullOrEmpty(kvp.Value.TenVattu2))
                {
                    string key2 = Helpers.NormalizeVietnameseString(kvp.Value.TenVattu2.Trim());
                    if (!string.IsNullOrEmpty(key2)) _lookupByTenPhu[key2] = kvp.Key;
                }
            }
            InitializeVatTuOptimization();
        }

        private void InitializeVatTuOptimization()
        {
            _optimizedVatTu = new Dictionary<string, (string, string, string, string, double, double)>();

            foreach (var item in vatTuLookup)
            {
                string ten1 = Helpers.NormalizeVietnameseString(item.Value.TenVattu);
                string ten2 = Helpers.NormalizeVietnameseString(item.Value.TenVattu2);
                string quyCach = RegexPatterns.QuantityUnit.Match(ten1).Value;

                _optimizedVatTu[item.Key] = (ten1, ten2, quyCach, item.Value.DonVi, item.Value.Dongia, item.Value.SoLuong);
            }
        }

        private void BuildIndexes()
        {
            if (_isIndexBuilt) return;

            _keywordIndex = new Dictionary<string, HashSet<string>>();
            _quyCachIndex = new Dictionary<string, HashSet<string>>();

            foreach (var kvp in _optimizedVatTu)
            {
                var words = kvp.Value.TenChuan.Split(new[] { ' ', '-', ',', ';', '/' },
                    StringSplitOptions.RemoveEmptyEntries).Where(w => w.Length >= 3);

                foreach (var word in words)
                {
                    if (!_keywordIndex.ContainsKey(word)) _keywordIndex[word] = new HashSet<string>();
                    _keywordIndex[word].Add(kvp.Key);
                }

                if (!string.IsNullOrEmpty(kvp.Value.QuyCach))
                {
                    if (!_quyCachIndex.ContainsKey(kvp.Value.QuyCach))
                        _quyCachIndex[kvp.Value.QuyCach] = new HashSet<string>();
                    _quyCachIndex[kvp.Value.QuyCach].Add(kvp.Key);
                }
            }
            _isIndexBuilt = true;
        }

        private string NormalizeNameForSearch(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            string result = Helpers.NormalizeVietnameseString(input).ToLower().Trim();

            // ✅ Sửa: thay ??= bằng if
            if (_synonymDictionary == null)
                _synonymDictionary = BuildSynonymDictionary();

            foreach (var synonym in _synonymDictionary)
            {
                if (result.Contains(synonym.Key))
                    result = result.Replace(synonym.Key, synonym.Value);
            }

            result = RegexPatterns.SpecialChars.Replace(result, " ");
            result = RegexPatterns.MultiWhitespace.Replace(result, " ").Trim();
            return result;
        }
        private Dictionary<string, string> BuildSynonymDictionary()
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["sài gòn"] = "saigon",
                ["sai gon"] = "saigon",
                ["cocacola"] = "coca cola",
                ["cô ca"] = "coca cola",
                ["cô ca cô la"] = "coca cola",
                ["pesi"] = "pepsi",
                ["redbull"] = "red bull",
            };
        }

        private static readonly (string Key, double Percent, string TenChuan, string QuyCach, string DonVi) NoMatch =
            (null, 0, null, null, null);

        private void Xulysohieuvattu(TbImportDetail tbImportDetail)
        {
            if (tbImportDetail == null || string.IsNullOrEmpty(tbImportDetail.Ten)) return;

            if (!_isIndexBuilt) BuildIndexes();

            if (_synonymDictionary == null)
                _synonymDictionary = BuildSynonymDictionary();

            string normalizedTen = NormalizeNameForSearch(tbImportDetail.Ten.Trim());
            string quyCach = RegexPatterns.QuantityUnit.Match(normalizedTen).Value;

            // 1. Exact match
            if (TryExactMatch(normalizedTen, quyCach, out var exact))
            {
                ApplyMatch(tbImportDetail, exact);
                return;
            }

            // 2. Tách từ khóa
            var words = normalizedTen.Split(new[] { ' ', '-', ',', ';', '/' },
                StringSplitOptions.RemoveEmptyEntries).Where(w => w.Length >= 3).ToList();

            var phrases = new List<string>();
            for (int i = 0; i < words.Count - 1; i++)
            {
                string phrase = words[i] + " " + words[i + 1];
                if (phrase.Length >= 5) phrases.Add(phrase);
            }

            // 3. Lọc ứng viên
            var candidates = GetCandidateKeys(words, phrases, quyCach);

            // 4. Tính điểm
            var best = FindBestMatch(candidates, normalizedTen, words, quyCach);

            if (best.HasValue) ApplyMatch(tbImportDetail, best.Value);
            else
            {
                tbImportDetail.SoHieu = GenerateResultString(normalizedTen);
                tbImportDetail.Percent = 0;
            }
        }

        private bool TryExactMatch(string normalizedTen, string quyCach,
            out (string Key, double Percent, string TenChuan, string QuyCach, string DonVi) result)
        {
            foreach (var kvp in _optimizedVatTu)
            {
                if (string.IsNullOrEmpty(quyCach) || kvp.Value.QuyCach == quyCach)
                {
                    if (NormalizeNameForSearch(kvp.Value.TenChuan) == normalizedTen
                        || NormalizeNameForSearch(kvp.Value.TenPhuChuan) == normalizedTen)
                    {
                        result = (kvp.Key, 100, kvp.Value.TenChuan, kvp.Value.QuyCach, kvp.Value.DonVi);
                        return true;
                    }
                }
            }
            result = NoMatch;
            return false;
        }

        private HashSet<string> GetCandidateKeys(List<string> words, List<string> phrases, string quyCach)
        {
            var set = new HashSet<string>();

            foreach (var w in words)
                if (_keywordIndex.TryGetValue(w, out var keys)) set.UnionWith(keys);

            foreach (var p in phrases)
                if (_keywordIndex.TryGetValue(p, out var keys)) set.UnionWith(keys);

            if (!string.IsNullOrEmpty(quyCach) && _quyCachIndex.TryGetValue(quyCach, out var qKeys))
                set.UnionWith(qKeys);

            if (set.Count == 0) set.UnionWith(_optimizedVatTu.Keys.Take(100));
            return set;
        }

        private (string Key, double Percent, string TenChuan, string QuyCach, string DonVi)? FindBestMatch(
            HashSet<string> candidates, string normalizedTen, List<string> words, string quyCach)
        {
            var results = new List<(string Key, double Percent, string TenChuan, string QuyCach, string DonVi, int MatchCount)>();
            double minPercent = 80;

            foreach (var key in candidates)
            {
                if (!_optimizedVatTu.TryGetValue(key, out var vt)) continue;

                string tenChuanHoa = NormalizeNameForSearch(vt.TenChuan);
                string tenKhongNgoac = RegexPatterns.BracketContent.Replace(tenChuanHoa, "").Trim();
                string tenHoaDonKhongNgoac = RegexPatterns.BracketContent.Replace(normalizedTen, "").Trim();

                int tokenNB = Fuzz.TokenSetRatio(tenKhongNgoac, tenHoaDonKhongNgoac);
                int partialNB = Fuzz.PartialRatio(tenKhongNgoac, tenHoaDonKhongNgoac);
                double pctNB = Math.Max(tokenNB, partialNB);

                int token = Fuzz.TokenSetRatio(tenChuanHoa, normalizedTen);
                int partial = Fuzz.PartialRatio(tenChuanHoa, normalizedTen);
                double pct = Math.Max(token, partial);

                double finalPct = Math.Max(pct, pctNB);
                int matchCount = words.Count(w => tenChuanHoa.Contains(w) || tenKhongNgoac.Contains(w));
                finalPct += matchCount * 5;

                string qcKho = vt.QuyCach?.Trim().ToLower() ?? "";
                if (!string.IsNullOrEmpty(quyCach) && !string.IsNullOrEmpty(qcKho)
                    && (qcKho == quyCach || qcKho.Contains(quyCach)))
                    finalPct += 20;

                if (pctNB > 80 && pctNB > pct && finalPct > 90) finalPct = 90;
                else if (pct > 80 && finalPct > 95) finalPct = 95;

                if (finalPct >= minPercent)
                    results.Add((key, Math.Min(finalPct, 100), vt.TenChuan, vt.QuyCach, vt.DonVi, matchCount));
            }

            if (results.Count == 0) return null;

            var best = results.OrderByDescending(r => r.MatchCount).ThenByDescending(r => r.Percent).First();
            return (best.Key, best.Percent, best.TenChuan, best.QuyCach, best.DonVi);
        }

        private static void ApplyMatch(TbImportDetail d,
            (string Key, double Percent, string TenChuan, string QuyCach, string DonVi) m)
        {
            d.SoHieu = m.Key;
            d.Percent = m.Percent;
            if (!string.IsNullOrEmpty(m.DonVi)) d.DVT = m.DonVi;
        }

        #endregion

        #region Load Cache Data

        private void LoadHoadonCT(string connectionst, TTinChung TTinChung)
        {
            try
            {
                string query = @"
                    SELECT hd.SoHD, hd.KyHieu, hd.NgayPH, hd.MaKhachHang, ct.NgayCT, ct.MaLoai
                    FROM Hoadon hd
                    INNER JOIN Chungtu ct ON hd.MaSo = ct.MaSo
                    WHERE hd.KyHieu <> '...'";

                var data = ExecuteQuery2(query, connectionst);
                lookupHoaDonCT.Clear();

                foreach (DataRow item in data.Rows)
                {
                    try
                    {
                        string soHD = RemoveLeadingZeros(item["SoHD"]?.ToString() ?? "").Trim();
                        string KyHieu = item["KyHieu"]?.ToString() ?? "";
                        DateTime ngayPH = ((DateTime)item["NgayCT"]).Date;
                        int Maloai = int.Parse(item["MaLoai"]?.ToString() ?? "0");
                        int type = Maloai == 8 ? 2 : 1;

                        lookupHoaDonCT.Add(("", soHD, KyHieu, ngayPH, type));
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Log($"Lỗi LoadHoadonCT: {ex.Message}"); }
            TTinChung.lookupHoaDonCT = lookupHoaDonCT;
        }

        private void Loadtbimport(string conectionst, TTinChung TTinChung)
        {
            try
            {
                tbimports = ExecuteQuery2("SELECT * FROM tbimport", conectionst);
                lookupTbImport.Clear();

                foreach (DataRow row in tbimports.Rows)
                {
                    try
                    {
                        string mst = row["Mst"]?.ToString() ?? "";
                        string shDon = RemoveLeadingZeros(row["SHDon"]?.ToString() ?? "").Trim();
                        DateTime nLap = row["NLap"] != DBNull.Value ? Convert.ToDateTime(row["NLap"]).Date : DateTime.MinValue;
                        int type = row["Type"] != DBNull.Value ? Convert.ToInt32(row["Type"]) : 0;

                        lookupTbImport.Add((mst, shDon, nLap, type));
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Log($"Lỗi Loadtbimport: {ex.Message}"); }
            TTinChung.lookupTbImport = lookupTbImport;
        }

        #endregion

        #region Button Run

        private async void btnRun_Click(object sender, EventArgs e)
        {
           // await TraCuuNCC(true, "3500405761", "AOEAN3B6ZH3B1FS", "13047", "D:\\DA3\\DA3\\Hoadon", 1);
             //await GetHoaDonMisa(true, "MVF0CBREBX2Q", "3502243033", "31768",1, "D:\\DA3\\DA3");  
            Xulydata();
            btnRun.Enabled = false;

            try
            {
                int totalRuns = int.Parse(txtsolanlap.Text);
                int totalLoops = int.Parse(txtSovongtai.Text);

                for (int loopCount = 1; loopCount <= totalLoops; loopCount++)
                {
                    Log($"🔄 ===== BẮT ĐẦU VÒNG LẶP {loopCount}/{totalLoops} =====");
                    labelControl3.Text = $"🔄 ===== BẮT ĐẦU VÒNG LẶP {loopCount}/{totalLoops} =====";

                    string computerName = Environment.MachineName;
                    tbCompany = ExecuteQuery("SELECT * FROM tbCompany WHERE IsRun = 1 and Saoviet = ? order by STT",
                        new OleDbParameter("?", computerName));

                    if (!tbCompany.Columns.Contains("RunCount"))
                        tbCompany.Columns.Add("RunCount", typeof(string));
                    if (!tbCompany.Columns.Contains("DateAccount"))
                        tbCompany.Columns.Add("DateAccount", typeof(DateTime));

                    foreach (DataRow row in tbCompany.Rows) row["RunCount"] = "0/0";

                    if (tbCompany.Rows.Count == 0)
                    {
                        Log($"⚠️ Vòng lặp {loopCount}: Không có công ty nào đang hoạt động!");
                        continue;
                    }

                    Log($"🚀 Vòng lặp {loopCount}: Xử lý {tbCompany.Rows.Count} công ty, mỗi công ty chạy {totalRuns} lần");

                    int maxParallel = int.Parse(txtSoluongtai.Text);
                    var semaphore = new SemaphoreSlim(maxParallel);
                    var companyQueue = new Queue<DataRow>();
                    foreach (DataRow item in tbCompany.Rows) companyQueue.Enqueue(item);

                    var tasks = new List<Task>();

                    for (int threadIndex = 0; threadIndex < maxParallel; threadIndex++)
                    {
                        tasks.Add(Task.Run(async () =>
                        {
                            while (true)
                            {
                                DataRow item = null;
                                lock (companyQueue)
                                {
                                    if (companyQueue.Count == 0) break;
                                    item = companyQueue.Dequeue();
                                }

                                string vbdbpath = item["Dbpath"]?.ToString() ?? "";
                                string companyName = item["Name"]?.ToString() ?? "Unknown";
                                string MST = item["MST"]?.ToString() ?? "";
                                int stt = int.Parse(item["STT"].ToString());

                                if (string.IsNullOrEmpty(vbdbpath))
                                {
                                    Log($"[{stt}] ⚠️ {companyName}: Không có Dbpath, bỏ qua!");
                                    continue;
                                }

                                await semaphore.WaitAsync();
                                try
                                {
                                    for (int runCount = 1; runCount <= totalRuns; runCount++)
                                    {
                                        if (!lstCongtyKhoa.Contains(MST))
                                        {
                                            UpdateRunCountOnUI(item, runCount, totalRuns);
                                            UpdateStatusOnUI(item, $"[{stt}] 🔄 {companyName} - Vòng {loopCount} - Lần {runCount}/{totalRuns}");
                                            Log($"[{stt}] 🔄 {companyName}: Vòng {loopCount} - Lần {runCount}/{totalRuns}");

                                            await TaihoadonCongty(vbdbpath, item);

                                            UpdateStatusOnUI(item, $"[{stt}] ✅ {companyName} - Vòng {loopCount} - Lần {runCount} - Hoàn thành");
                                        }
                                    }

                                    Log($"[{stt}] 📄 {companyName}: Bắt đầu xử lý XML...");
                                    UpdateStatusOnUI(item, $"[{stt}] 📄 {companyName} - Đang xử lý XML...");
                                    UpdateRunCountOnUI(item, totalRuns, totalRuns);

                                    try
                                    {
                                        string connectionString2 = AppConstants.BuildConnectionString(vbdbpath);
                                        Xulytooltrunggian(connectionString2);
                                        Log($"[{stt}] ✅ {companyName}: Hoàn thành xử lý XML");
                                        UpdateStatusOnUI(item, $"[{stt}] ✅ {companyName} - Hoàn thành XML");
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"[{stt}] ❌ Lỗi XML {companyName}: {ex.Message}");
                                        UpdateStatusOnUI(item, $"[{stt}] ❌ {companyName} - Lỗi XML: {ex.Message}");
                                    }
                                }
                                finally
                                {
                                    semaphore.Release();
                                }
                            }
                        }));
                    }

                    await Task.WhenAll(tasks);
                    Log($"✅ Vòng lặp {loopCount}/{totalLoops} hoàn thành!");

                    if (loopCount < totalLoops)
                    {
                        Log($"⏳ Chờ 5s trước vòng lặp tiếp theo...");
                        await Task.Delay(5000);
                    }
                }

                Log($"✅ ===== HOÀN THÀNH TẤT CẢ {totalLoops} VÒNG LẶP =====");
                this.Close();
            }
            catch (Exception ex) { Log($"❌ Lỗi: {ex.Message}"); }
            finally { btnRun.Enabled = true; }
        }

        private void Xulytooltrunggian(string connectionString2)
        {
            var getResgistry = ExecuteQuery2("SELECT * FROM tbRegister", connectionString2);

            ExecuteQueryResult2("UPDATE tbRegister SET IsRunning = ?", connectionString2,
                new OleDbParameter("?", "1"));

            string hoadonpath = getResgistry.Rows[0]["Hoadonpath"].ToString();
            string statusFilePath = Path.Combine(hoadonpath, "status.txt");
            File.WriteAllText(statusFilePath, "1");

            string backPath = Directory.GetParent(hoadonpath)?.FullName ?? hoadonpath;
            string compind = Path.Combine(backPath, "Tools\\Debug\\SaovietTax.exe");

            if (File.Exists(compind))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo { FileName = compind, UseShellExecute = true };
                Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show($"Không tìm thấy file: {compind}");
            }
        }

        #endregion

        #region DocfileXmlOne

        private bool Loaiborow(string name) => name.ToLower().Contains("điều chỉnh");

        private string CapitalizeFirstLetters(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            return char.ToUpper(input[0]) + input.Substring(1);
        }

        private void GetMST(string connectionString2)
        {
            var kq = ExecuteQuery2("SELECT * FROM License", connectionString2, null);
            if (kq.Rows.Count > 0)
            {
                MSTCongTY = kq.Rows[0]["MaSoThue"].ToString();
                CCCD = kq.Rows[0]["CCCD"].ToString();
            }
        }

        private bool CheckExistKH(string mst)
        {
            if (!string.IsNullOrEmpty(mst))
            {
                if (tbKhachhang.AsEnumerable().Any(row =>
                    row.Field<string>("MST") == mst || row.Field<string>("SoHieu") == mst))
                    return true;
            }
            return false;
        }

        private bool Kiemtrahoadon(string SHDon, DateTime NLap, string MST, int type, TTinChung TTinChung)
        {
            return TTinChung.lookupTbImport.Contains((MST, SHDon, NLap, type));
        }

        private bool KiemtrahoadonCT(string SoHD, string KyHieu, DateTime NLap, string Mst, int tpye, TTinChung TTinChung)
        {
            if (Mst == "KL") Mst = "00";
            if (Mst.Length < 10)
                return TTinChung.lookupHoaDonCT.Any(m => m.SoHD == SoHD && m.KyHieu == KyHieu && m.NLap == NLap && m.Type == type);
            return TTinChung.lookupHoaDonCT.Contains((Mst, SoHD, KyHieu, NLap, tpye));
        }

        public void InitCustomer(int Maphanloai, string Sohieu, string Ten, string Diachi, string Mst,
            string cccd, string sdt, string connectionst, TTinChung TTinChung)
        {
            if (string.IsNullOrEmpty(sdt)) sdt = "xxx";
            string diachiKHVni = !string.IsNullOrEmpty(Diachi)
                ? Helpers.ConvertUnicodeToVni(Diachi)
                : Helpers.ConvertUnicodeToVni("Bổ sung địa chỉ");

            if (string.IsNullOrEmpty(Mst))
            {
                if (string.IsNullOrEmpty(cccd))
                {
                    Sohieu = GenerateAbbreviation(Helpers.ConvertVniToUnicode(Ten),
                        TTinChung.tbKhachhang.AsEnumerable().Select(row => row.Field<string>("SoHieu")).ToList()).ToUpper();
                    csohieu = Sohieu;
                    Mst = "00";

                    int suffix = 1;
                    string originalSohieu = Sohieu;
                    while (TTinChung.tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                    {
                        Sohieu = $"{originalSohieu}_{suffix}";
                        suffix++;
                    }
                }
                else { Sohieu = cccd.Substring(cccd.Length - 6); Mst = cccd; }
            }
            else
            {
                if (string.IsNullOrEmpty(Sohieu))
                {
                    Sohieu = Helpers.GetLastFourDigits(Mst.Replace("-", ""));
                    if (TTinChung.tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                        Sohieu = "0" + Sohieu;
                    if (TTinChung.tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                        Sohieu = "00" + Sohieu;
                }
            }

            if (TTinChung.tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                Sohieu = Sohieu + "_1";

            if (Mst == Sohieu && Mst.Length <= 8) Mst = "00";
            else
            {
                if (Mst.Length > 8)
                {
                    Sohieu = Helpers.GetLastFourDigits(Mst.Replace("-", ""));
                    if (TTinChung.tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                        Sohieu = "0" + Sohieu;
                    if (TTinChung.tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                        Sohieu = "00" + Sohieu;
                }
            }

            string query;
            OleDbParameter[] parameters;

            if (mstcongty == "3501972322")
            {
                var getlastid = ExecuteQuery("SELECT MAX(MaSo) FROM KhachHang").Rows[0]["Expr1000"].ToString();
                query = "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu,Ten,DiaChi,MST,Tel) VALUES (?,?,?,?,?,?,?)";
                parameters = new OleDbParameter[]
                {
                    new OleDbParameter("?", int.Parse(getlastid) + 1),
                    new OleDbParameter("?", Maphanloai),
                    new OleDbParameter("?", Sohieu),
                    new OleDbParameter("?", Ten),
                    new OleDbParameter("?", diachiKHVni),
                    new OleDbParameter("?", Mst),
                    new OleDbParameter("?", sdt),
                };
            }
            else
            {
                query = "INSERT INTO KhachHang (MaPhanLoai,SoHieu,Ten,DiaChi,MST,Tel) VALUES (?,?,?,?,?,?)";
                parameters = new OleDbParameter[]
                {
                    new OleDbParameter("?", Maphanloai),
                    new OleDbParameter("?", Sohieu),
                    new OleDbParameter("?", Ten),
                    new OleDbParameter("?", diachiKHVni),
                    new OleDbParameter("?", Mst),
                    new OleDbParameter("?", sdt),
                };
            }

            try
            {
                ExecuteQueryResult2(query, connectionst, parameters);
                if (mstcongty != "3501972322")
                {
                    TTinChung.tbKhachhang = ExecuteQuery2("SELECT * FROM KhachHang", connectionst);
                }
            }
            catch (Exception ex) { XtraMessageBox.Show(ex.Message + "  " + Ten + "  " + cccd); }
        }

        #endregion

        #region Schedule & Startup

        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit1.Checked) AddToStartup();
            else RemoveFromStartup();
        }

        public void AddToStartup()
        {
            try
            {
                string exePath = Application.ExecutablePath;
                if (!File.Exists(exePath))
                {
                    XtraMessageBox.Show($"❌ File EXE không tồn tại:\n{exePath}");
                    return;
                }

                using (RegistryKey rk = Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (rk == null) { XtraMessageBox.Show("❌ Không thể mở Registry!"); return; }

                    string existingValue = rk.GetValue(AppConstants.StartupAppName) as string;
                    string exeWithArgs = $"\"{exePath}\" -autostart";

                    if (!string.IsNullOrEmpty(existingValue) && existingValue == exeWithArgs) return;

                    if (rk.GetValue(AppConstants.StartupAppName) != null)
                        rk.DeleteValue(AppConstants.StartupAppName);

                    rk.SetValue(AppConstants.StartupAppName, exeWithArgs);

                    UpdateAutoTaiSetting();
                    XtraMessageBox.Show($"✅ Đã đăng ký Startup!\n\nTên: {AppConstants.StartupAppName}\nĐường dẫn: {exeWithArgs}");
                }
            }
            catch (Exception ex) { XtraMessageBox.Show($"❌ Lỗi: {ex.Message}"); }
        }

        private void UpdateAutoTaiSetting()
        {
            try
            {
                string computerName = Environment.MachineName;
                string checkQuery = "SELECT COUNT(*) FROM tbsetting WHERE Saoviet = ?";
                DataTable dt = ExecuteQuery(checkQuery, new OleDbParameter("?", computerName));
                int count = Convert.ToInt32(dt.Rows[0][0]);

                if (count > 0)
                {
                    ExecuteQueryResult("UPDATE tbsetting SET AutoTai = ? WHERE Saoviet = ?",
                        new OleDbParameter("?", 1),
                        new OleDbParameter("?", computerName));
                }
                else
                {
                    ExecuteQueryResult("INSERT INTO tbsetting (Saoviet, AutoTai,Soluongtai,Timeout) VALUES (?, ?,?,?)",
                        new OleDbParameter("?", computerName),
                        new OleDbParameter("?", 1),
                        new OleDbParameter("?", txtSoluongtai.Text),
                        new OleDbParameter("?", txttimeout.Text));
                }
            }
            catch (Exception ex) { Log($"❌ Lỗi SetAutoTaiForComputer: {ex.Message}"); }
        }

        public void RemoveFromStartup()
        {
            try
            {
                using (RegistryKey rk = Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (rk?.GetValue(AppConstants.StartupAppName) != null)
                        rk.DeleteValue(AppConstants.StartupAppName, false);
                }
            }
            catch { }
        }

        public static void DeleteSchedule(string taskName)
        {
            try
            {
                string cmd = $"schtasks /delete /tn \"{taskName}\" /f";
                Process.Start("cmd", "/c " + cmd).WaitForExit();
                XtraMessageBox.Show($"✅ Đã xóa lịch '{taskName}'!");
            }
            catch (Exception ex) { XtraMessageBox.Show($"❌ Lỗi xóa: {ex.Message}"); }
        }

        public static void CreateSchedule(string taskName, int hour, int minute = 0)
        {
            try
            {
                string exePath = Application.ExecutablePath;
                string timeStr = $"{hour:D2}:{minute:D2}";

                RunSchTasks($"/delete /tn \"{taskName}\" /f", false);

                string tr = $"\\\"{exePath}\\\" -autostart";
                string args = $"/create /tn \"{taskName}\" /tr \"{tr}\" /sc daily /st {timeStr} /f";

                RunSchTasks(args, true);
                XtraMessageBox.Show($"Đã tạo lịch '{taskName}' lúc {timeStr} mỗi ngày.");
            }
            catch (Exception ex) { XtraMessageBox.Show(ex.ToString()); }
        }

        private static void RunSchTasks(string arguments, bool throwIfError)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "schtasks.exe",
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (Process p = Process.Start(psi))
            {
                string output = p.StandardOutput.ReadToEnd();
                string error = p.StandardError.ReadToEnd();
                p.WaitForExit();

                if (p.ExitCode != 0 && throwIfError)
                    throw new Exception($"ExitCode: {p.ExitCode}\r\nOutput:\r\n{output}\r\nError:\r\n{error}");
            }
        }

        #endregion

        #region Event Handlers

        private void textEdit4_EditValueChanged(object sender, EventArgs e)
        {
            string computerName = Environment.MachineName;
            ExecuteQueryResult("UPDATE tbsetting SET Soluongtai = ? where Saoviet=?",
                new OleDbParameter("?", txtSoluongtai.Text),
                new OleDbParameter("?", computerName));
        }

        private void txttimeout_EditValueChanged(object sender, EventArgs e)
        {
            string computerName = Environment.MachineName;
            ExecuteQueryResult("UPDATE tbsetting SET Timeout = ? where Saoviet=?",
                new OleDbParameter("?", txttimeout.Text),
                new OleDbParameter("?", computerName));
        }

        private void chkMoc1_CheckedChanged(object sender, EventArgs e)
        {
            txtBlock1.Enabled = chkMoc1.Checked;
            if (!chkMoc1.Checked) txtBlock1.Text = "";
        }

        private void chkMoc2_CheckedChanged(object sender, EventArgs e)
        {
            txtBlock2.Enabled = chkMoc2.Checked;
            if (!chkMoc2.Checked) txtBlock2.EditValue = "";
        }

        private void chkMoc3_CheckedChanged(object sender, EventArgs e)
        {
            txtBlock3.Enabled = chkMoc3.Checked;
            if (!chkMoc3.Checked) txtBlock3.EditValue = "";
        }

        private void txtBlock1_EditValueChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBlock1.Text))
            {
                string computerName = Environment.MachineName;
                ExecuteQueryResult("UPDATE tbsetting SET Block1 = ? where Saoviet=?",
                    new OleDbParameter("?", txtBlock1.Text),
                    new OleDbParameter("?", computerName));
                DeleteSchedule(AppConstants.TaskBlock1);
            }
        }

        private void txtBlock2_EditValueChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBlock2.Text))
            {
                string computerName = Environment.MachineName;
                ExecuteQueryResult("UPDATE tbsetting SET Block2 = ? where Saoviet=?",
                    new OleDbParameter("?", txtBlock2.Text),
                    new OleDbParameter("?", computerName));
                DeleteSchedule(AppConstants.TaskBlock2);
            }
        }

        private void txtBlock3_EditValueChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBlock3.Text))
            {
                string computerName = Environment.MachineName;
                ExecuteQueryResult("UPDATE tbsetting SET Block3 = ? where Saoviet=?",
                    new OleDbParameter("?", txtBlock3.Text),
                    new OleDbParameter("?", computerName));
                DeleteSchedule(AppConstants.TaskBlock3);
            }
        }

        private void SaveBlockSetting(string blockNumber, string value)
        {
            try
            {
                string computerName = Environment.MachineName;
                string checkQuery = "SELECT COUNT(*) FROM tbsetting WHERE Saoviet = ?";
                DataTable dt = ExecuteQuery(checkQuery, new OleDbParameter("?", computerName));
                int count = Convert.ToInt32(dt.Rows[0][0]);

                if (count > 0)
                {
                    ExecuteQueryResult($"UPDATE tbsetting SET {blockNumber} = ? WHERE Saoviet = ?",
                        new OleDbParameter("?", value),
                        new OleDbParameter("?", computerName));
                }
                else
                {
                    ExecuteQueryResult($"INSERT INTO tbsetting (Saoviet, {blockNumber},Soluongtai,Timeout) VALUES (?, ?,?,?)",
                        new OleDbParameter("?", computerName),
                        new OleDbParameter("?", value),
                        new OleDbParameter("?", txtSoluongtai.Text),
                        new OleDbParameter("?", txttimeout.Text));
                }
            }
            catch (Exception ex) { Log($"❌ Lỗi SaveBlockSetting: {ex.Message}"); }
        }

        private void CreateScheduleFromText(string taskName, string timeText)
        {
            if (DateTime.TryParse(timeText, out DateTime time))
            {
                try { CreateSchedule(taskName, time.Hour, time.Minute); }
                catch (Exception ex) { XtraMessageBox.Show(ex.Message); }
            }
        }

        private void txtBlock1_Validated(object sender, EventArgs e)
        {
            SaveBlockSetting("Block1", txtBlock1.Text);
            CreateScheduleFromText(AppConstants.TaskBlock1, txtBlock1.Text);
        }

        private void txtBlock2_Validated(object sender, EventArgs e)
        {
            SaveBlockSetting("Block2", txtBlock2.Text);
            CreateScheduleFromText(AppConstants.TaskBlock2, txtBlock2.Text);
        }

        private void txtBlock3_Validated(object sender, EventArgs e)
        {
            SaveBlockSetting("Block3", txtBlock3.Text);
            CreateScheduleFromText(AppConstants.TaskBlock3, txtBlock3.Text);
        }

        private void timeEdit1_EditValueChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(timeEdit1.Text))
            {
                string computerName = Environment.MachineName;
                ExecuteQueryResult("UPDATE tbsetting SET Block1 = ? where Saoviet=?",
                    new OleDbParameter("?", timeEdit1.Text),
                    new OleDbParameter("?", computerName));
                DeleteSchedule(AppConstants.TaskBlock1);
            }
        }

        private void timeEdit1_Validated(object sender, EventArgs e)
        {
            SaveBlockSetting("Block1", timeEdit1.Text);
            CreateScheduleFromText(AppConstants.TaskBlock1, timeEdit1.Text);
        }

        private void timeEdit2_EditValueChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(timeEdit2.Text))
            {
                string computerName = Environment.MachineName;
                ExecuteQueryResult("UPDATE tbsetting SET Block2 = ? where Saoviet=?",
                    new OleDbParameter("?", timeEdit2.Text),
                    new OleDbParameter("?", computerName));
                DeleteSchedule(AppConstants.TaskBlock2);
            }
        }

        private void timeEdit2_Validated(object sender, EventArgs e)
        {
            SaveBlockSetting("Block2", timeEdit2.Text);
            CreateScheduleFromText(AppConstants.TaskBlock2, timeEdit2.Text);
        }

        private void timeEdit3_EditValueChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(timeEdit3.Text))
            {
                string computerName = Environment.MachineName;
                ExecuteQueryResult("UPDATE tbsetting SET Block3 = ? where Saoviet=?",
                    new OleDbParameter("?", timeEdit3.Text),
                    new OleDbParameter("?", computerName));
                DeleteSchedule(AppConstants.TaskBlock3);
            }
        }

        private void timeEdit3_Validated(object sender, EventArgs e)
        {
            SaveBlockSetting("Block3", timeEdit3.Text);
            CreateScheduleFromText(AppConstants.TaskBlock3, timeEdit3.Text);
        }

        private void comboBoxEdit1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string computerName = Environment.MachineName;
            int selectedMonth = DateTime.Now.Month;

            if (comboBoxEdit1.SelectedItem != null && comboBoxEdit1.SelectedItem.ToString() != "")
            {
                if (int.Parse(comboBoxEdit1.EditValue.ToString()) < DateTime.Now.Month)
                    selectedMonth = int.Parse(comboBoxEdit1.EditValue.ToString());
                else
                    selectedMonth = Convert.ToInt32(comboBoxEdit1.SelectedItem);

                thangchon = selectedMonth;
            }

            try
            {
                DataTable dt = ExecuteQuery("SELECT COUNT(*) FROM tbRemember WHERE Saoviet = ?",
                    new OleDbParameter("?", computerName));
                int count = Convert.ToInt32(dt.Rows[0][0]);

                if (count > 0)
                {
                    if (dateEdit1.DateTime != DateTime.MinValue)
                    {
                        ExecuteQueryResult("UPDATE tbRemember SET ThangCT = ?, TuNgay =?, DenNgay= ? WHERE Saoviet = ?",
                            new OleDbParameter("?", selectedMonth),
                            new OleDbParameter("?", dateEdit1.DateTime),
                            new OleDbParameter("?", dateEdit2.DateTime),
                            new OleDbParameter("?", computerName));
                    }
                }
                else
                {
                    ExecuteQueryResult("INSERT INTO tbRemember (Saoviet, ThangCT) VALUES (?, ?)",
                        new OleDbParameter("?", computerName),
                        new OleDbParameter("?", selectedMonth));
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi cập nhật ThangCT: {ex.Message}");
                MessageBox.Show($"Lỗi: {ex.Message}");
            }
            Kichhoatngay();
        }

        private void dateEdit1_EditValueChanged(object sender, EventArgs e)
        {
            DateTime selectedDate = dateEdit1.DateTime;
            comboBoxEdit1.EditValue = selectedDate.Month;

            DateTime lastDay = new DateTime(selectedDate.Year, selectedDate.Month,
                DateTime.DaysInMonth(selectedDate.Year, selectedDate.Month));

            if (lastDay.Month != DateTime.Now.Month)
                dateEdit2.DateTime = lastDay;
            else
            {
                if (dateEdit1.DateTime > dateEdit2.DateTime)
                {
                    if (dateEdit1.DateTime < DateTime.Now) dateEdit2.DateTime = DateTime.Now.Date;
                    else dateEdit2.DateTime = dateEdit1.DateTime;
                }
            }
            capnhatngay();
        }

        private void capnhatngay()
        {
            string computerName = Environment.MachineName;
            ExecuteQueryResult("UPDATE tbRemember SET TuNgay = ?,DenNgay=? WHERE Saoviet = ?",
                new OleDbParameter("?", dateEdit1.DateTime),
                new OleDbParameter("?", dateEdit2.DateTime),
                new OleDbParameter("?", computerName));
        }

        private void dateEdit2_EditValueChanged(object sender, EventArgs e) => capnhatngay();

        private void labelControl1_Click(object sender, EventArgs e) { }
        private void gridControl1_Click(object sender, EventArgs e) { }
        private void panel1_Paint(object sender, PaintEventArgs e) { }
        private void panel1_Paint_1(object sender, PaintEventArgs e) { }
        private void gridView1_RowStyle(object sender, RowStyleEventArgs e) { }

        #endregion

        #region Save Invoices Bulk

        private async Task SaveAllInvoicesBulk(List<TbImport> invoices, int type, string connectionString)
        {
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                await conn.OpenAsync();
                using (OleDbTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        string sqlParent = @"INSERT INTO tbImport (SHDon, KHHDon, NLap, Ten, Noidung, TKNo, TKCo, TkThue, Mst, [Status], Ngaytao, TongTien, Vat, TPhi, TgTCThue, TgTThue, [Type], InvoiceType, IsHaschild, TVat, Vat2, TVat2, Vat3, TVat3, TgTCThue1, TgTCThue2, TgTCThue3, Khmshdon, hdon, [Path],IsMD) 
                                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

                        string sqlDetail = @"INSERT INTO tbimportdetail (ParentId, SoHieu, SoLuong, DonGia, DVT, Ten, MaCT, TKNo, TKCo, TTien, [Percent], Tchat,SoPSGoc,VAT) 
                                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

                        foreach (var item in invoices)
                        {
                            using (OleDbCommand cmdParent = new OleDbCommand(sqlParent, conn, trans))
                            {
                                cmdParent.Parameters.AddWithValue("?", item.SHDon ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.KHHDon ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.NLap);
                                cmdParent.Parameters.AddWithValue("?", item.Ten ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.Noidung ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.TKNo ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.TKCo ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.TkThue);
                                cmdParent.Parameters.AddWithValue("?", item.Mst ?? "");
                                cmdParent.Parameters.AddWithValue("?", "0");
                                cmdParent.Parameters.AddWithValue("?", DateTime.Now.ToShortDateString());
                                cmdParent.Parameters.AddWithValue("?", item.TongTien);
                                cmdParent.Parameters.AddWithValue("?", item.Vat);
                                cmdParent.Parameters.AddWithValue("?", item.TPhi);
                                cmdParent.Parameters.AddWithValue("?", Math.Round(item.TgTCThue));
                                cmdParent.Parameters.AddWithValue("?", Math.Round(item.TgTThue));
                                cmdParent.Parameters.AddWithValue("?", type);
                                cmdParent.Parameters.AddWithValue("?", "0");
                                cmdParent.Parameters.AddWithValue("?", item.IsHaschild == null ? "1" : item.IsHaschild);
                                cmdParent.Parameters.AddWithValue("?", item.TVat);
                                cmdParent.Parameters.AddWithValue("?", item.Vat2 ?? "0");
                                cmdParent.Parameters.AddWithValue("?", item.TVat2);
                                cmdParent.Parameters.AddWithValue("?", item.Vat3 ?? "0");
                                cmdParent.Parameters.AddWithValue("?", item.TVat3);
                                cmdParent.Parameters.AddWithValue("?", item.TgTCThue1);
                                cmdParent.Parameters.AddWithValue("?", item.TgTCThue2);
                                cmdParent.Parameters.AddWithValue("?", item.TgTCThue3);
                                cmdParent.Parameters.AddWithValue("?", item.Khmshdon ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.hdon ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.Path ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.IsMD);
                                await cmdParent.ExecuteNonQueryAsync();
                            }

                            int parentID;
                            using (OleDbCommand cmdId = new OleDbCommand("SELECT @@IDENTITY", conn, trans))
                                parentID = Convert.ToInt32(await cmdId.ExecuteScalarAsync());

                            foreach (var dt in item.tbImportDetails)
                            {
                                using (OleDbCommand cmdDetail = new OleDbCommand(sqlDetail, conn, trans))
                                {
                                    cmdDetail.Parameters.AddWithValue("?", parentID);
                                    cmdDetail.Parameters.AddWithValue("?", dt.SoHieu ?? "");
                                    cmdDetail.Parameters.AddWithValue("?", dt.Soluong);
                                    cmdDetail.Parameters.AddWithValue("?", dt.Dongia);
                                    cmdDetail.Parameters.AddWithValue("?", dt.DVT ?? "");
                                    cmdDetail.Parameters.AddWithValue("?", dt.Ten ?? "");
                                    cmdDetail.Parameters.AddWithValue("?", "");
                                    cmdDetail.Parameters.AddWithValue("?", dt.TKNo ?? "");
                                    cmdDetail.Parameters.AddWithValue("?", dt.TKCo ?? "");
                                    cmdDetail.Parameters.AddWithValue("?", dt.TTien);
                                    cmdDetail.Parameters.AddWithValue("?", dt.Percent);
                                    cmdDetail.Parameters.AddWithValue("?", dt.Tchat);
                                    cmdDetail.Parameters.AddWithValue("?", dt.TTien);
                                    cmdDetail.Parameters.AddWithValue("?", dt.Vat);
                                    await cmdDetail.ExecuteNonQueryAsync();
                                }
                            }
                        }

                        trans.Commit();
                        Log($"Đã lưu thành công {invoices.Count} hóa đơn vào Database!");
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        XtraMessageBox.Show("Lỗi lưu hàng loạt: " + ex.Message, "Lỗi Database");
                    }
                }
            }
        }

        #endregion

        #region Xuly711

        public void Xuly711(TbImport fileImport)
        {
            var getlist711 = fileImport.tbImportDetails.Where(m => m.TKCo == "711").ToList();

            if (getlist711.Count > 0 && tbLicense.Rows[0].Field<string>("col711") == "1")
            {
                if (getlist711.Count == 1)
                {
                    var get711 = getlist711.FirstOrDefault();
                    if (get711 != null)
                    {
                        var remainlist = fileImport.tbImportDetails
                            .Where(m => m != get711 && !string.IsNullOrEmpty(m.DVT) && m.Dongia != 0).ToList();
                        int index = 0;
                        double totalth = remainlist.Sum(m => m.TTien);
                        double total711 = get711.TTien;

                        foreach (var it2 in remainlist)
                        {
                            if (index < remainlist.Count - 1)
                                it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                            else
                            {
                                it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                                double reTotal = Math.Round(fileImport.TgTCThue - remainlist.Sum(m => m.TTien));
                                if (reTotal > 0) it2.TTien += reTotal;
                                if (reTotal == -1) it2.TTien -= 1;
                            }
                            index += 1;
                        }
                        fileImport.tbImportDetails.Remove(get711);
                    }
                }
                else
                {
                    var get771s = getlist711.ToList();
                    foreach (var i7 in get771s) fileImport.tbImportDetails.Remove(i7);

                    var remainlist = fileImport.tbImportDetails
                        .Where(m => m.TKCo != "711" && !string.IsNullOrEmpty(m.DVT) && m.Dongia != 0).ToList();
                    double totalth = remainlist.Sum(m => m.TTien);
                    double total711 = get771s.Sum(m => m.TTien);
                    if (total711 < 0) total711 = -total711;

                    int index = 0;
                    foreach (var it2 in remainlist)
                    {
                        if (remainlist.Count > 1)
                        {
                            if (index < remainlist.Count - 1)
                                it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                            else
                            {
                                it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                                double reTotal = Math.Round(fileImport.TgTCThue - remainlist.Sum(m => m.TTien));
                                it2.TTien += reTotal;
                            }
                        }
                        else it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                        index += 1;
                    }
                }
            }
            else
            {
                if (getlist711.Count > 1)
                {
                    var getfirst = getlist711.FirstOrDefault();
                    var lstremain = getlist711.Skip(1).ToList();
                    getfirst.TTien += lstremain.Sum(m => m.TTien);
                    if (getfirst.TTien < 0) getfirst.TTien = -getfirst.TTien;
                    foreach (var it in lstremain) fileImport.tbImportDetails.Remove(it);
                }
            }
        }

        #endregion

        #region Unused Methods (giữ lại để tránh breaking change)

        [Obsolete("Dùng GettokkenAsync thay thế")]
        public void Gettokken(string username, string password, ref string currentToken, string connectist, DataRow rowCopy)
        {
            var outcome = GettokkenAsync(username, password, rowCopy).GetAwaiter().GetResult();
            currentToken = outcome.Result == LoginResult.WrongCredentials ? "-1" : outcome.Token;
        }

        #endregion

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                _pdfCts?.Cancel();
                _pdfSignal?.Release(1000);
                // Chờ tối đa 10s để worker dừng
                _pdfWorkerTask?.Wait(10000);
            }
            catch { }
            base.OnFormClosing(e);
        }
    }

    #region Extension Methods

    public static class HttpClientExtensions
    {
        public static async Task<string> GetStringAsync(this Form1.CompanyHttpClient client, string url)
        {
            var response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
    }



    #region Tải hoá đơn gốc
  
    #endregion


    #endregion




}