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
            public string Mstnb { get; set; }
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
            public int IsDownload { get; set; }
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

            var checkEdit2 = new RepositoryItemCheckEdit();
            checkEdit2.ValueChecked = 1;
            checkEdit2.ValueUnchecked = 0;
            gridView1.Columns["IsDownload"].ColumnEdit = checkEdit2;
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
            var isdownload=companyRow?["IsDownload"]?.ToString() == "1";
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

                            bool ok = await DownloadSingleInvoiceAsync(client, capturedInvoice, companyName, cts.Token, isdownload);

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
            string companyName, CancellationToken ct, bool isdownload)
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
                        if (isdownload)
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
                if (NCC == "VIETTEL" || NCC == "MISA" || NCC == "FAST" || NCC == "BKAV" || NCC == "MINVOICE" || NCC == "VNPT" || NCC == "Softdream" || NCC == "Win Tech" || NCC == "Cyber" || NCC == "Visnam")
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

            if (task.PathXml.Contains("HDRa"))
            {

            }

            if (task.Ncc == "MISA")
            {
               await GetHoaDonMisa(false, task.Mabimat, task.Mst, task.Shdon, task.Type, savedPath);
            }
            else if (task.Ncc == "VIETTEL")
            {
             await TraCuuNCC(false, task.Mst, task.Mabimat, task.Shdon, savedPath, task.Type);
            }
            else if (task.Ncc == "FAST")
            {
               await TaiFAST(false, task.Mst, task.Shdon,task.Khhd, task.Mabimat, savedPath, task.Type);
            }
            else if (task.Ncc == "BKAV")
            {
               await TaiBKAV(false, task.Mst, task.Shdon, task.Khhd, task.Mabimat, savedPath, task.Type);
            }
            else if (task.Ncc == "MINVOICE")
            {
               await TaiMINVOICE(false, task.Mst, task.Shdon, task.Khhd, task.Mabimat,savedPath, task.Type);
            }
            else if (task.Ncc == "VNPT")
            {
                if (task.Mst.StartsWith("0300555450"))
                {
                    await TaiHDVNPTPETROLIMEX(false, task.Mst, task.Shdon, task.Khhd, task.Mabimat, savedPath, task.Type);
                }
                else
                {
                    await TaiHDVNPT(false, task.Mst, task.Shdon, task.Khhd, task.Mabimat, savedPath, task.Type);
                }
                   
            }
            else if (task.Ncc == "Softdream")
            {
                await TaiSoftdream(false, task.Mst, task.Shdon, task.Khhd, task.Mabimat, savedPath, task.Type);
            }
            else if (task.Ncc == "Win Tech")
            {
                await TaiWinTech(false, task.Mst, task.Shdon, task.Khhd, task.Mabimat, savedPath, task.Type);
            }
            else if (task.Ncc == "Cyber")
            {
                await TaiCyber(false, task.Mst, task.Shdon, task.Khhd, task.Mabimat, savedPath, task.Type);
            }
            else if (task.Ncc == "Visnam")
            {
                await TaiVisnam(false, task.Mst, task.Shdon, task.Khhd, task.Mabimat, savedPath, task.Type);
            }
        }
        public async Task TaiVisnam(bool toggle, string mst, string shdon, string KHHDon, string mabimat, string savedPath, int type)
        {
            string url = "https://tracuu.vin-hoadon.com/tracuuhoadon/tracuuxacthuc/tracuuhd";
            string downloadPath = "";

            if (type == 1)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDVao\\" + dateEdit1.DateTime.Month;
            if (type == 2)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDRa\\" + dateEdit1.DateTime.Month;

            Directory.CreateDirectory(downloadPath);

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless=new");
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
                                "FileSystemAccessAPI,IdleDetection,SyncDisabled,BackForwardCache");
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

            // uBlock (nếu có)
            string ublockPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Extensions", "uBlock");
            if (Directory.Exists(ublockPath))
                options.AddArgument("--load-extension=" + ublockPath);

            // Download
            options.AddUserProfilePreference("download.default_directory", downloadPath);
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
            options.AddUserProfilePreference("profile.default_content_setting_values.clipboard", 2);
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddUserProfilePreference("profile.password_manager_leak_detection", false);

            IWebDriver driver = null;
            try
            {
                string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                service.HideCommandPromptWindow = true;

                driver = new ChromeDriver(service, options);
                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                driver.Navigate().GoToUrl(url);

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

                // ═══════════════════════════════════════════════════════════
                // HÀM ĐIỀN MÃ SỐ BÍ MẬT — GỌI LẠI MỖI LẦN THỬ CAPTCHA
                // ═══════════════════════════════════════════════════════════
                Action dienMaSoBiMat = () =>
                {
                    try
                    {
                        bool filled = (bool)js.ExecuteScript(@"
                    var input = document.querySelector('input[name=""MaSoBiMat""]') || 
                                document.querySelector('input[placeholder=""Mã số bí mật""]');
                    if (input) {
                        input.value = arguments[0];
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.dispatchEvent(new Event('change', { bubbles: true }));
                        input.dispatchEvent(new Event('blur', { bubbles: true }));
                        return true;
                    }
                    return false;
                ", mabimat);

                        if (filled)
                            Console.WriteLine("✅ Đã điền Mã số bí mật");
                        else
                            Console.WriteLine("⚠ Không tìm thấy input MaSoBiMat");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("⚠ Lỗi điền MaSoBiMat: " + ex.Message);
                    }
                };

                // ============================================================
                // BƯỚC 1: ĐIỀN MÃ BÍ MẬT LẦN ĐẦU
                // ============================================================
                await Task.Delay(2000);
                dienMaSoBiMat();

                // ============================================================
                // BƯỚC 2: OCR CAPTCHA (có retry)
                // ============================================================
                string result = string.Empty;
                int maxRetry = 3;

                for (int attempt = 1; attempt <= maxRetry; attempt++)
                {
                    try
                    {
                        Console.WriteLine($"🔍 Lần thử captcha {attempt}/{maxRetry}");

                        // ✅ QUAN TRỌNG: Điền lại MaSoBiMat mỗi lần thử (từ lần 2)
                        if (attempt > 1)
                        {
                            await Task.Delay(300);
                            dienMaSoBiMat();
                        }

                        // Tìm ảnh captcha
                        IWebElement captchaImg = wait.Until(d => d.FindElement(By.XPath("//img[@alt='Captcha']")));

                        string base64 = (string)js.ExecuteScript(@"
                    var canvas = document.createElement('canvas');
                    var img = arguments[0];
                    canvas.width = img.naturalWidth || img.width;
                    canvas.height = img.naturalHeight || img.height;
                    var ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0);
                    return canvas.toDataURL('image/png').substring(22);
                ", captchaImg);

                        byte[] captchaBytes = Convert.FromBase64String(base64);

                        string rawPath = Path.Combine(downloadPath, $"captcha_raw_{attempt}.png");
                        File.WriteAllBytes(rawPath, captchaBytes);

                        result = Kingsofthp.OcrCaptcha(captchaBytes);
                        Console.WriteLine($"   OCR → '{result}'");

                        if (string.IsNullOrEmpty(result) || result.Length < 3 || result.Length > 6)
                        {
                            Console.WriteLine($"   ⚠ Kết quả không hợp lệ, reload...");
                            Kingsofthp.ReloadCaptcha(driver, wait);
                            await Task.Delay(300);
                            continue;
                        }

                        // Điền captcha
                        IWebElement inputCaptcha = wait.Until(d => d.FindElement(By.Name("curentcap")));
                        inputCaptcha.Clear();
                        inputCaptcha.SendKeys(result);

                        js.ExecuteScript(@"
                    var input = arguments[0];
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                ", inputCaptcha);

                        await Task.Delay(500);

                        // Submit form
                        js.ExecuteScript(@"
                    var form = document.querySelector('form.m-login__form');
                    if (form) {
                        var submitBtn = form.querySelector('button[type=""submit""]');
                        if (submitBtn) {
                            submitBtn.click();
                        } else {
                            form.submit();
                        }
                    }
                ");
                        await Task.Delay(1500);

                        // Kiểm tra captcha sai
                        bool captchaWrong = false;
                        try
                        {
                            var errorElements = driver.FindElements(By.XPath(
                                "//*[contains(text(),'mã xác thực') or contains(text(),'Mã xác thực')]"));
                            foreach (var el in errorElements)
                            {
                                string txt = el.Text;
                                if (txt.Contains("không") || txt.Contains("sai"))
                                {
                                    captchaWrong = true;
                                    break;
                                }
                            }
                        }
                        catch { }

                        if (captchaWrong)
                        {
                            Console.WriteLine("   ❌ Captcha sai, reload...");
                            Kingsofthp.ReloadCaptcha(driver, wait);
                            // Vòng lặp sẽ tự điền lại MaSoBiMat ở đầu lần sau
                            continue;
                        }

                        Console.WriteLine("   ✅ Captcha đúng, chờ modal xuất hiện");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Lỗi lần {attempt}: {ex.Message}");
                        try { Kingsofthp.ReloadCaptcha(driver, wait); } catch { }
                        await Task.Delay(500);
                    }
                }

                if (string.IsNullOrEmpty(result))
                {
                    XtraMessageBox.Show("Không thể giải captcha sau nhiều lần thử.");
                    return;
                }

                // ============================================================
                // BƯỚC 3: CHỜ MODAL VÀ CLICK NÚT TẢI PDF
                // ============================================================
                Console.WriteLine("   ⏳ Chờ modal hiển thị nút 'Tải PDF'...");

                try
                {
                    var waitPopup = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                    IWebElement btnDownloadPdf = waitPopup.Until(d =>
                    {
                        try
                        {
                            var btn = d.FindElement(By.XPath(
                                "//div[contains(@class, 'modal-content')]//button[contains(@class, 'btn-warning') and .//span[normalize-space()='Tải PDF']]"));
                            return (btn != null && btn.Displayed) ? btn : null;
                        }
                        catch (NoSuchElementException)
                        {
                            return null;
                        }
                    });

                    Console.WriteLine("   ✅ Đã tìm thấy nút 'Tải PDF'");

                    js.ExecuteScript("arguments[0].scrollIntoView({block: 'center'});", btnDownloadPdf);
                    await Task.Delay(300);

                    try
                    {
                        btnDownloadPdf.Click();
                        Console.WriteLine("   ✅ Đã click nút tải PDF (native)");
                    }
                    catch (ElementClickInterceptedException)
                    {
                        js.ExecuteScript("arguments[0].click();", btnDownloadPdf);
                        Console.WriteLine("   ✅ Đã click nút tải PDF (JS)");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("   ❌ Không tìm thấy nút 'Tải PDF' trong modal sau 20s");
                    throw;
                }

                // ============================================================
                // BƯỚC 4: CHỜ FILE PDF TẢI XONG VÀ ĐỔI TÊN
                // ============================================================
                await Task.Delay(1000);

                string pdfFile = WaitForPdf(shdon, mst, downloadPath, 15);
                if (pdfFile == null)
                {
                    XtraMessageBox.Show("Không thể tải file PDF");
                    return;
                }
                await Task.Delay(500);

                if (pdfFile != null && File.Exists(pdfFile))
                {
                    string fileMoi = Path.Combine(
                        downloadPath,
                        mst + "_" + shdon + "_" + KHHDon + ".pdf"
                    );

                    if (!string.Equals(pdfFile, fileMoi, StringComparison.OrdinalIgnoreCase))
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
                            if (File.Exists(pdfFile))
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
                Console.WriteLine("❌ Lỗi TaiVisnam: " + ex.Message);

                try
                {
                    if (driver != null)
                    {
                        var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                        screenshot.SaveAsFile(Path.Combine(downloadPath, "debug_error.png"));
                        File.WriteAllText(Path.Combine(downloadPath, "debug_error.html"), driver.PageSource);
                        Console.WriteLine("📸 Đã lưu debug_error.png + debug_error.html");
                    }
                }
                catch { }

                throw;
            }
            finally
            {
                try { if (driver != null) driver.Quit(); } catch { }
                try { if (driver != null) driver.Dispose(); } catch { }
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
        }
        public async Task TaiCyber(bool toggle, string mst, string shdon, string KHHDon, string mabimat, string savedPath, int type)
        {
            string url = "https://tracuuhoadon2.xcyber.vn/#/tracuuhoadon/tracuu";
            string downloadPath = "";

            if (type == 1)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDVao\\" + dateEdit1.DateTime.Month;
            if (type == 2)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDRa\\" + dateEdit1.DateTime.Month;

            Directory.CreateDirectory(downloadPath);

            ChromeOptions options = new ChromeOptions();
           // options.AddArgument("--headless=new");
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
                                "FileSystemAccessAPI,IdleDetection,SyncDisabled,BackForwardCache");
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

            // uBlock (nếu có)
            string ublockPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Extensions", "uBlock");
            if (Directory.Exists(ublockPath))
                options.AddArgument("--load-extension=" + ublockPath);

            // Download
            options.AddUserProfilePreference("download.default_directory", downloadPath);
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
            options.AddUserProfilePreference("profile.default_content_setting_values.clipboard", 2);
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddUserProfilePreference("profile.password_manager_leak_detection", false);

            IWebDriver driver = null;
            try
            {
                string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                service.HideCommandPromptWindow = true; // Ẩn cửa sổ cmd đen

                driver = new ChromeDriver(service, options);

                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                driver.Navigate().GoToUrl(url);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

                IWebElement inputMaSoBiMat = wait.Until(d =>
                {
                    try
                    {
                        var el = d.FindElement(By.Name("MaSoBiMat"));
                        return el.Displayed && el.Enabled ? el : null;
                    }
                    catch
                    {
                        return null;
                    }
                });

                inputMaSoBiMat.Clear();
                inputMaSoBiMat.SendKeys(mabimat);

                // BƯỚC 1: LẤY CHUỖI BASE64 TỪ TRÌNH DUYỆT
                string base64Data = wait.Until(d =>
                {
                    try
                    {
                        IWebElement img = d.FindElement(By.CssSelector("img[alt='Captcha']"));
                        string src = img.GetAttribute("src");
                        if (!string.IsNullOrEmpty(src) && src.StartsWith("data:image"))
                        {
                            return src;
                        }
                        return null;
                    }
                    catch (NoSuchElementException)
                    {
                        return null;
                    }
                });

                // BƯỚC 2: GỌI PYTHON OCR (DDDDOCR) THAY CHO TESSERACT
                string result = string.Empty;
                if (!string.IsNullOrEmpty(base64Data))
                {
                    result = Kingsofthp.OcrCaptcha(Convert.FromBase64String(base64Data.Split(',')[1]));
                }

                Console.WriteLine($"KẾT QUẢ CUỐI CÙNG (ddddocr): {result}");

                // BƯỚC 3: ĐIỀN VÀ BẤM NÚT "TRA CỨU"
                if (!string.IsNullOrEmpty(result))
                {
                    try
                    {
                        IWebElement inputCaptcha = wait.Until(d => d.FindElement(By.Name("curentcap")));
                        inputCaptcha.Clear();

                        // Gõ từng phím để giả lập thao tác người dùng (giúp web Angular kích hoạt nút)
                        foreach (char c in result)
                        {
                            inputCaptcha.SendKeys(c.ToString());
                            System.Threading.Thread.Sleep(50);
                        }

                        IWebElement btnSubmit = wait.Until(d =>
                            d.FindElement(By.XPath("//button[contains(., 'Tra cứu')]"))
                        );

                        wait.Until(d => btnSubmit.Enabled);
                        btnSubmit.Click();

                        Console.WriteLine("Đã nhấn Tra cứu thành công!");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Lỗi khi điền hoặc click: {ex.Message}");
                    }
                }

                // ═══════════════════════════════════════════════════════
                // BƯỚC 4: CHỜ POPUP HÓA ĐƠN HIỆN LÊN
                // ═══════════════════════════════════════════════════════

                IWebElement popup = wait.Until(d =>
                {
                    try
                    {
                        var el = d.FindElement(By.CssSelector("div.modal.fade.in.show #modal-dialog"));
                        return el.Displayed ? el : null;
                    }
                    catch (NoSuchElementException)
                    {
                        return null;
                    }
                });

                Console.WriteLine("✅ Popup hóa đơn đã hiện lên!");

                // Chờ thêm để nội dung hóa đơn render xong
                System.Threading.Thread.Sleep(1500);

                // ═══════════════════════════════════════════════════════
                // BƯỚC 5: SCROLL XUỐNG NÚT "Tải PDF" VÀ BẤM
                // ═══════════════════════════════════════════════════════

                IWebElement btnTaiPDF = wait.Until(d =>
                {
                    try
                    {
                        var buttons = d.FindElements(By.CssSelector("#footer-btn button"));
                        foreach (var btn in buttons)
                        {
                            if (btn.Text.Trim().Contains("Tải PDF"))
                            {
                                return btn.Displayed && btn.Enabled ? btn : null;
                            }
                        }
                        return null;
                    }
                    catch
                    {
                        return null;
                    }
                });

                // Scroll xuống nút Tải PDF
                ((IJavaScriptExecutor)driver).ExecuteScript(
                    "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                    btnTaiPDF
                );

                System.Threading.Thread.Sleep(800); // chờ scroll xong

                // Bấm nút Tải PDF
                try
                {
                    btnTaiPDF.Click();
                }
                catch (ElementClickInterceptedException)
                {
                    // Fallback: dùng JS click nếu bị overlay che
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", btnTaiPDF);
                }
                Console.WriteLine("✅ Đã bấm nút Tải PDF!");

                // ═══════════════════════════════════════════════════════
                // BƯỚC 6: CHỜ PDF TẢI XONG
                // ═══════════════════════════════════════════════════════

                string pdfFile = WaitForPdf(shdon, mst, downloadPath, 20);
                if (pdfFile == null)
                {
                    XtraMessageBox.Show("Không thể tải file PDF");
                    return;
                }
                Thread.Sleep(500);

                // ═══════════════════════════════════════════════════════
                // BƯỚC 7: ĐỔI TÊN FILE
                // ═══════════════════════════════════════════════════════

                if (pdfFile != null && File.Exists(pdfFile))
                {
                    string fileMoi = Path.Combine(
                        downloadPath,
                        mst + "_" + shdon + "_" + KHHDon + ".pdf"
                    );

                    if (!string.Equals(pdfFile, fileMoi, StringComparison.OrdinalIgnoreCase))
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
                            if (File.Exists(pdfFile))
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
                Console.WriteLine("❌ Lỗi TaiCyber: " + ex.Message);

                // Debug: lưu ảnh + HTML
                try
                {
                    if (driver != null)
                    {
                        var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                        screenshot.SaveAsFile(Path.Combine(downloadPath, "debug_error.png"));
                        File.WriteAllText(Path.Combine(downloadPath, "debug_error.html"), driver.PageSource);
                        Console.WriteLine("📸 Đã lưu debug_error.png + debug_error.html");
                    }
                }
                catch { }

                throw;
            }
            finally
            {
                try { if (driver != null) driver.Quit(); } catch { }
                try { if (driver != null) driver.Dispose(); } catch { }
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
        }
        public async Task TaiWinTech(bool toggle, string mst, string shdon, string KHHDon, string mabimat, string savedPath, int type)
        {
            string url = "https://tracuu.wininvoice.vn/bill";
            string downloadPath = "";

            if (type==1)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDVao\\" + dateEdit1.DateTime.Month;
            if (type==2)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDRa\\" + dateEdit1.DateTime.Month;

            Directory.CreateDirectory(downloadPath);

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless=new");
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
                                "FileSystemAccessAPI,IdleDetection,SyncDisabled,BackForwardCache");
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

            // uBlock (nếu có)
            string ublockPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Extensions", "uBlock");
            if (Directory.Exists(ublockPath))
                options.AddArgument("--load-extension=" + ublockPath);

            // Download
            options.AddUserProfilePreference("download.default_directory", downloadPath);
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
            options.AddUserProfilePreference("profile.default_content_setting_values.clipboard", 2);
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddUserProfilePreference("profile.password_manager_leak_detection", false);

            IWebDriver driver = null;
            try
            {
                string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                service.HideCommandPromptWindow = true; // Ẩn cửa sổ cmd đen

                driver = new ChromeDriver(service, options);

                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                driver.Navigate().GoToUrl(url);

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                // ═══════════════════════════════════════════════════════
                // BƯỚC 1: ĐIỀN FORM TRA CỨU
                // ═══════════════════════════════════════════════════════

                wait.Until(d => d.FindElement(By.Name("private_code")));

                var txtPrivateCode = driver.FindElement(By.Name("private_code"));
                txtPrivateCode.Clear();
                txtPrivateCode.SendKeys(mabimat);

                var txtCmnpKey = driver.FindElement(By.Name("cmpn_key"));
                txtCmnpKey.Clear();
                txtCmnpKey.SendKeys(mst);

                var btnSubmit = driver.FindElement(By.CssSelector("button.btn.blue[type='submit']"));
                btnSubmit.Click();

                Console.WriteLine($"✅ Đã submit form tra cứu: MST={mst}, SoHD={mabimat}");

                // ═══════════════════════════════════════════════════════
                // BƯỚC 2: CHỜ TRANG KẾT QUẢ /bill/inv
                // ═══════════════════════════════════════════════════════

                wait.Until(d => d.Url.Contains("/bill/inv"));
                Console.WriteLine($"✅ Đã chuyển sang trang kết quả: {driver.Url}");

                // ═══════════════════════════════════════════════════════
                // BƯỚC 3: TÌM VÀ BẤM NÚT DOWNLOAD
                // ═══════════════════════════════════════════════════════

                // Có thể có nhiều dòng trong bảng (nhiều hóa đơn)
                // → Chọn dòng có "Số HĐ" khớp với shdon
                IWebElement btnDownload = null;

                try
                {
                    // Cách 1: Tìm nút download có href chứa "download=1"
                    btnDownload = wait.Until(d =>
                    {
                        var buttons = d.FindElements(By.CssSelector("a[href*='download=1']"));
                        return buttons.Count > 0 ? buttons[0] : null;
                    });

                    Console.WriteLine($"✅ Tìm thấy nút download: {btnDownload.GetAttribute("href")}");
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("⚠ Không tìm thấy nút download, thử cách khác...");

                    // Cách 2: Tìm theo class + title
                    try
                    {
                        btnDownload = driver.FindElement(
                            By.CssSelector("a.btn-info[title='Tải file hóa đơn']"));
                    }
                    catch
                    {
                        // Cách 3: Tìm theo icon download
                        try
                        {
                            btnDownload = driver.FindElement(
                                By.XPath("//a[contains(@class, 'btn-info') and .//i[contains(@class, 'fa-download')]]"));
                        }
                        catch
                        {
                            // Cách 4: Lấy nút download cuối cùng trong bảng
                            var allBtns = driver.FindElements(
                                By.XPath("//td[contains(@class, 'td-actions')]//a[contains(@class, 'btn-info')]"));
                            if (allBtns.Count > 0)
                                btnDownload = allBtns[allBtns.Count - 1];
                        }
                    }
                }

                if (btnDownload == null)
                {
                    XtraMessageBox.Show("Không tìm thấy nút tải hóa đơn");
                    return;
                }

                // Bấm nút download
                btnDownload.Click();
                Console.WriteLine("📥 Đã bấm nút download, chờ file...");

                // ═══════════════════════════════════════════════════════
                // BƯỚC 4: CHỜ PDF TẢI XONG
                // ═══════════════════════════════════════════════════════

                string pdfFile = WaitForZip(shdon, downloadPath, 15);
                if (pdfFile == null)
                {
                    XtraMessageBox.Show("Không thể tải file PDF");
                    return;
                }
                Thread.Sleep(500);

                // ═══════════════════════════════════════════════════════
                // BƯỚC 5: ĐỔI TÊN FILE
                // ═══════════════════════════════════════════════════════

                if (pdfFile != null && File.Exists(pdfFile))
                {
                    string fileMoi = Path.Combine(
                        downloadPath,
                        mst + "_" + shdon + "_" + KHHDon + ".pdf"
                    );

                    if (!string.Equals(pdfFile, fileMoi, StringComparison.OrdinalIgnoreCase))
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
                            if (File.Exists(pdfFile))
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
                Console.WriteLine("❌ Lỗi TaiWinTech: " + ex.Message);

                // Debug: lưu ảnh + HTML
                try
                {
                    if (driver != null)
                    {
                        var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                        screenshot.SaveAsFile(Path.Combine(downloadPath, "debug_error.png"));
                        File.WriteAllText(Path.Combine(downloadPath, "debug_error.html"), driver.PageSource);
                        Console.WriteLine("📸 Đã lưu debug_error.png + debug_error.html");
                    }
                }
                catch { }

                throw;
            }
            finally
            {
                try { if (driver != null) driver.Quit(); } catch { }
                try { if (driver != null) driver.Dispose(); } catch { }
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
        }
        public async Task TaiSoftdream(bool toggle, string mst, string shdon, string KHHDon, string mabimat, string savedPath, int type)
        {
            string url = "https://tracuu.easyinvoice.vn/Search/Index";
            string downloadPath = "";

            if (type==1)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDVao\\" + dateEdit1.DateTime.Month;
            if (type==2)
                downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDRa\\" + dateEdit1.DateTime.Month;

            Directory.CreateDirectory(downloadPath);

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless=new");
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
                                "FileSystemAccessAPI,IdleDetection,SyncDisabled,BackForwardCache");
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

            // uBlock (nếu có)
            string ublockPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Extensions", "uBlock");
            if (Directory.Exists(ublockPath))
                options.AddArgument("--load-extension=" + ublockPath);

            // Download
            options.AddUserProfilePreference("download.default_directory", downloadPath);
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
            options.AddUserProfilePreference("profile.default_content_setting_values.clipboard", 2);
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddUserProfilePreference("profile.password_manager_leak_detection", false);

            IWebDriver driver = null;
            try
            {
                string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                service.HideCommandPromptWindow = true; // Ẩn cửa sổ cmd đen

                driver = new ChromeDriver(service, options);

                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                driver.Navigate().GoToUrl(url);

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                // ============================================================
                // BƯỚC 1: ĐIỀN MST + MÃ BÍ MẬT
                // ============================================================
                try
                {
                    var txtTaxCode = wait.Until(d => d.FindElement(By.Id("taxCode")));
                    txtTaxCode.Clear();
                    txtTaxCode.SendKeys(mst);

                    var txtFKey = wait.Until(d => d.FindElement(By.Id("iFkey")));
                    txtFKey.Clear();
                    txtFKey.SendKeys(mabimat);

                    Console.WriteLine("✅ Đã điền MST + FKey");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("⚠ Không tìm thấy ô nhập taxCode/iFkey: " + ex.Message);
                }

                // ============================================================
                // BƯỚC 2: OCR CAPTCHA
                // ============================================================
                string result = string.Empty;
                int maxRetry = 1;

                for (int attempt = 1; attempt <= maxRetry; attempt++)
                {
                    try
                    {
                        Console.WriteLine($"🔍 Lần thử captcha {attempt}/{maxRetry}");

                        IWebElement captchaImg = wait.Until(d => d.FindElement(By.Id("captcha")));

                        string base64 = (string)((IJavaScriptExecutor)driver).ExecuteScript(@"
                    var canvas = document.createElement('canvas');
                    var img = arguments[0];
                    canvas.width = img.naturalWidth || img.width;
                    canvas.height = img.naturalHeight || img.height;
                    var ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0);
                    return canvas.toDataURL('image/png').substring(22);
                ", captchaImg);

                        byte[] captchaBytes = Convert.FromBase64String(base64);

                        string rawPath = Path.Combine(downloadPath, $"captcha_raw_{attempt}.png");
                        //rawPath = "C:\\Users\\Admin\\Desktop\\captcha_raw_1.png";
                        File.WriteAllBytes(rawPath, captchaBytes);

                        result = Kingsofthp.OcrCaptcha(captchaBytes);
                        Console.WriteLine($"   OCR → '{result}'");

                        if (string.IsNullOrEmpty(result) || result.Length < 3 || result.Length > 5)
                        {
                            Console.WriteLine($"   ⚠ Kết quả không hợp lệ, reload...");
                            Kingsofthp.ReloadCaptcha(driver, wait);
                            System.Threading.Thread.Sleep(300);
                            continue;
                        }

                        IWebElement inputCaptcha = wait.Until(d => d.FindElement(By.Id("Capcha")));
                        inputCaptcha.Clear();
                        inputCaptcha.SendKeys(result);

                        IWebElement btnSubmit = wait.Until(d =>
                            d.FindElement(By.XPath("//button[contains(., 'Tra cứu')] | //input[@value='Tra cứu']")));

                        wait.Until(d => btnSubmit.Enabled);
                        btnSubmit.Click();

                        Console.WriteLine("   ✅ Đã nhấn Tra cứu");

                        System.Threading.Thread.Sleep(1500);

                        bool captchaWrong = false;
                        try
                        {
                            var errorElements = driver.FindElements(By.XPath(
                                "//*[contains(text(),'mã xác thực') or contains(text(),'Mã xác thực')]"));
                            foreach (var el in errorElements)
                            {
                                string txt = el.Text;
                                if (txt.Contains("không") || txt.Contains("sai"))
                                {
                                    captchaWrong = true;
                                    break;
                                }
                            }
                        }
                        catch { }

                        if (captchaWrong)
                        {
                            Console.WriteLine("   ❌ Captcha sai, reload...");
                            Kingsofthp.ReloadCaptcha(driver, wait);
                            continue;
                        }

                        Console.WriteLine("   ✅ Captcha đúng, chờ popup");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Lỗi lần {attempt}: {ex.Message}");
                        try { Kingsofthp.ReloadCaptcha(driver, wait); } catch { }
                        System.Threading.Thread.Sleep(500);
                    }
                }

                if (string.IsNullOrEmpty(result))
                {
                    XtraMessageBox.Show("Không thể giải captcha sau nhiều lần thử.");
                    return;
                }

                // ============================================================
                // BƯỚC 3: CHỜ POPUP VÀ CLICK NÚT TẢI PDF
                // ============================================================
                Console.WriteLine("   ⏳ Chờ popup và click nút tải PDF...");

                try
                {
                    var waitPopup = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                    IWebElement btnDownloadPdf = waitPopup.Until(d =>
                    {
                        try
                        {
                            var btn = d.FindElement(By.Name("downloadPdfAndFileAttach"));
                            return (btn != null && btn.Displayed) ? btn : null;
                        }
                        catch (NoSuchElementException)
                        {
                            return null;
                        }
                    });

                    Console.WriteLine("   ✅ Đã tìm thấy nút 'Tải PDF & đính kèm'");

                    ((IJavaScriptExecutor)driver).ExecuteScript(
                        "arguments[0].scrollIntoView({block: 'center'});",
                        btnDownloadPdf);

                    System.Threading.Thread.Sleep(300);

                    try
                    {
                        btnDownloadPdf.Click();
                        Console.WriteLine("   ✅ Đã click nút (native)");
                    }
                    catch (ElementClickInterceptedException)
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", btnDownloadPdf);
                        Console.WriteLine("   ✅ Đã click nút (JS)");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("   ❌ Không tìm thấy nút 'Tải PDF' sau 20s");
                    try
                    {
                        var btn = driver.FindElement(By.CssSelector("button[name='downloadPdfAndFileAttach']"));
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", btn);
                        Console.WriteLine("   ✅ Đã click (CSS fallback)");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ❌ Fallback lỗi: {ex.Message}");
                    }
                }

                // ============================================================
                // BƯỚC 4: CHỜ FILE ZIP TẢI XONG
                // ============================================================
                System.Threading.Thread.Sleep(800);

                string zipFile = WaitForZip(shdon, downloadPath, 30);
                //if (zipFile == null)
                //{
                //    XtraMessageBox.Show("Không thể tải file ZIP");
                //    return;
                //}

                Console.WriteLine($"✅ Đã tải ZIP: {zipFile}");

                // ============================================================
                // BƯỚC 5: GIẢI NÉN ZIP → LẤY PDF → XÓA ZIP + FOLDER TẠM
                // ============================================================
                string pdfFile = WaitForPdf(shdon, mst, downloadPath, 15);
                if (pdfFile == null)
                {
                    XtraMessageBox.Show("Không thể tải file PDF");
                    return;
                }
                Thread.Sleep(500);

                // ═══════════════════════════════════════════════════════
                // BƯỚC 5: ĐỔI TÊN FILE
                // ═══════════════════════════════════════════════════════

                if (pdfFile != null && File.Exists(pdfFile))
                {
                    string fileMoi = Path.Combine(
                        downloadPath,
                        mst + "_" + shdon + "_" + KHHDon + ".pdf"
                    );

                    if (!string.Equals(pdfFile, fileMoi, StringComparison.OrdinalIgnoreCase))
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
                            if (File.Exists(pdfFile))
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
                Console.WriteLine("❌ Lỗi TaiSoftdream: " + ex.Message);

                try
                {
                    if (driver != null)
                    {
                        var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                        screenshot.SaveAsFile(Path.Combine(downloadPath, "debug_error.png"));
                        File.WriteAllText(Path.Combine(downloadPath, "debug_error.html"), driver.PageSource);
                        Console.WriteLine("📸 Đã lưu debug_error.png + debug_error.html");
                    }
                }
                catch { }

                throw;
            }
            finally
            {
                try { if (driver != null) driver.Quit(); } catch { }
                try { if (driver != null) driver.Dispose(); } catch { }
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
        }
        private string WaitForZip(string shdon, string folder, int timeoutSeconds)
        {
            DateTime start = DateTime.Now;
            string zipFile = null;

            // ═══════════════════════════════════════════════════════
            // BƯỚC 1: CHỜ FILE ZIP
            // ═══════════════════════════════════════════════════════

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

                    // Lấy tất cả file ZIP
                    string[] zipFiles = Directory.GetFiles(folder, "*.zip");
                    if (zipFiles.Length == 0)
                    {
                        Thread.Sleep(500);
                        continue;
                    }

                    // Ưu tiên 1: Lọc file ZIP có tên chứa shdon
                    var matched = zipFiles
                        .Where(f => Path.GetFileName(f)
                            .IndexOf(shdon, StringComparison.OrdinalIgnoreCase) >= 0)
                        .ToArray();

                    // Fallback: Không khớp tên → lấy tất cả ZIP
                    var candidates = matched.Length > 0 ? matched : zipFiles;

                    // Chọn file mới nhất
                    string newestFile = null;
                    DateTime newestTime = DateTime.MinValue;

                    foreach (string file in candidates)
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
                    {
                        Console.WriteLine($"✅ Tìm thấy ZIP: {Path.GetFileName(newestFile)} ({size1} bytes)");
                        zipFile = newestFile;
                        break;
                    }

                    // File đang ghi tiếp → chờ
                    Thread.Sleep(500);
                }
                catch (IOException)
                {
                    Thread.Sleep(500);
                }
                catch (UnauthorizedAccessException)
                {
                    Thread.Sleep(500);
                }

                Thread.Sleep(500);
            }

            // Timeout → log debug
            if (zipFile == null)
            {
                Console.WriteLine($"⏱ Hết {timeoutSeconds}s không tìm thấy ZIP chứa '{shdon}' trong: {folder}");
                try
                {
                    var debugFiles = Directory.GetFiles(folder);
                    Console.WriteLine($"📁 Có {debugFiles.Length} file trong thư mục:");
                    foreach (var f in debugFiles)
                    {
                        var fi = new FileInfo(f);
                        Console.WriteLine($"   - {fi.Name} ({fi.Length} bytes)");
                    }
                }
                catch { }
                return null;
            }

            // ═══════════════════════════════════════════════════════
            // BƯỚC 2: GIẢI NÉN → LẤY PDF RA NGOÀI
            // ═══════════════════════════════════════════════════════

            try
            {
                // Thư mục tạm để giải nén
                string tempFolder = Path.Combine(folder, "_temp_extract");
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);

                Directory.CreateDirectory(tempFolder);

                // Giải nén
                ZipFile.ExtractToDirectory(zipFile, tempFolder);
                Console.WriteLine($"✅ Đã giải nén: {zipFile}");

                // Tìm tất cả file PDF trong thư mục giải nén (bao gồm subfolder)
                var pdfFiles = Directory.GetFiles(tempFolder, "*.pdf", SearchOption.AllDirectories);

                if (pdfFiles.Length == 0)
                {
                    Console.WriteLine("⚠ Không tìm thấy file PDF nào trong ZIP");
                    Directory.Delete(tempFolder, true);
                    return null;
                }

                // Di chuyển tất cả PDF ra ngoài folder chính
                string firstPdf = null;
                foreach (var pdf in pdfFiles)
                {
                    string destPath = Path.Combine(folder, Path.GetFileName(pdf));

                    // Nếu file đích đã tồn tại → xóa trước
                    if (File.Exists(destPath))
                        File.Delete(destPath);

                    // Di chuyển
                    File.Move(pdf, destPath);
                    Console.WriteLine($"📄 Đã lấy PDF ra: {Path.GetFileName(destPath)}");

                    if (firstPdf == null)
                        firstPdf = destPath;
                }

                // Xóa thư mục tạm
                try { Directory.Delete(tempFolder, true); } catch { }

                // ═══════════════════════════════════════════════════════
                // BƯỚC 3: XÓA FILE ZIP
                // ═══════════════════════════════════════════════════════

                try
                {
                    // Đợi 1 chút để chắc chắn file không còn bị lock
                    Thread.Sleep(200);

                    if (File.Exists(zipFile))
                    {
                        File.Delete(zipFile);
                        Console.WriteLine($"🗑 Đã xóa ZIP: {Path.GetFileName(zipFile)}");
                    }
                }
                catch (IOException ioEx)
                {
                    Console.WriteLine($"⚠ Không xóa được ZIP (đang bị lock): {ioEx.Message}");
                }
                catch (UnauthorizedAccessException uaEx)
                {
                    Console.WriteLine($"⚠ Không có quyền xóa ZIP: {uaEx.Message}");
                }

                return firstPdf;   // Trả về PDF đầu tiên
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi giải nén: {ex.Message}");

                // Fallback: thử xóa thư mục tạm
                try
                {
                    string tempFolder = Path.Combine(folder, "_temp_extract");
                    if (Directory.Exists(tempFolder))
                        Directory.Delete(tempFolder, true);
                }
                catch { }

                return null;
            }
        }

        public async Task TaiHDVNPT(
      bool toggle,
      string mst,
      string shdon,
      string KHHD,
      string mabm,
      string savedPath,
      int type)
        {
            string downloadPath = "";

            if (type == 1)
            {
                downloadPath = Path.Combine(
                    savedPath,
                    $"HD{dateEdit1.DateTime.Year}",
                    "HDVao",
                    dateEdit1.DateTime.Month.ToString());
            }
            else if (type == 2)
            {
                downloadPath = Path.Combine(
                    savedPath,
                    $"HD{dateEdit1.DateTime.Year}",
                    "HDRa",
                    dateEdit1.DateTime.Month.ToString());
            }

            if (string.IsNullOrEmpty(downloadPath))
            {
                XtraMessageBox.Show("Chưa chọn đầu vào hoặc đầu ra.");
                return;
            }

            Directory.CreateDirectory(downloadPath);

            await Task.Run(async () =>
            {
                // ═══════════════════════════════════════════════════════
                // CHROME OPTIONS — CHỐNG BOT + CHỐNG POPUP DOWNLOAD
                // ═══════════════════════════════════════════════════════
                ChromeOptions options = new ChromeOptions();

                options.AddArgument("--headless=new");
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

                options.AddUserProfilePreference("download.default_directory", downloadPath);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("safebrowsing.enabled", true);
                options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
                options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                options.AddUserProfilePreference("profile.default_content_setting_values.popups", 1);

                options.AddUserProfilePreference("profile.default_content_setting_values.media_stream", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.notifications", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.clipboard", 2);
                options.AddUserProfilePreference("credentials_enable_service", false);
                options.AddUserProfilePreference("profile.password_manager_enabled", false);
                options.AddUserProfilePreference("profile.password_manager_leak_detection", false);

                IWebDriver driver = null;
                try
                {
                    string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                    service.HideCommandPromptWindow = true;

                    driver = new ChromeDriver(service, options);
                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);

                    // ═══════════════════════════════════════════════════════
                    // 🛡️ CDP: DOWNLOAD BEHAVIOR + BYPASS CSP
                    // ═══════════════════════════════════════════════════════
                    try
                    {
                        if (driver is ChromeDriver chromeDriver)
                        {
                            chromeDriver.ExecuteCdpCommand(
                                "Page.setDownloadBehavior",
                                new Dictionary<string, object>
                                {
                            { "behavior", "allow" },
                            { "downloadPath", downloadPath }
                                });

                            chromeDriver.ExecuteCdpCommand("Page.enable", new Dictionary<string, object>());
                            chromeDriver.ExecuteCdpCommand(
                                "Page.setBypassCSP",
                                new Dictionary<string, object> { { "enabled", true } });

                            Console.WriteLine("✅ CDP download behavior đã bật");
                        }
                    }
                    catch (Exception cdpEx)
                    {
                        Console.WriteLine($"⚠️ CDP setup: {cdpEx.Message}");
                    }

                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

                    // ═══════════════════════════════════════════════════════
                    // 🌐 TẠO URL ĐỘNG TỪ MST NGƯỜI BÁN
                    // ═══════════════════════════════════════════════════════
                    // MST có thể có dạng:
                    //   - "3600244645"      (10 số)
                    //   - "3600244645-009"  (13 số có gạch)
                    // VNPT cấp subdomain theo MST bỏ gạch
                    string mstFull = mst.Replace("-", "").Replace(".", "").Replace(" ", "").Trim();
                    string mstGoc = mst.Contains("-") ? mst.Split('-')[0].Trim() : mstFull;

                    // Thử lần lượt các URL
                    string[] candidateUrls = new[]
                    {
                $"https://{mstFull}-tt78.vnpt-invoice.com.vn/",   // MST đầy đủ (13 số)
                $"https://{mstGoc}-tt78.vnpt-invoice.com.vn/"     // MST gốc (10 số)
            };

                    Console.WriteLine("🌐 Các URL sẽ thử:");
                    foreach (var u in candidateUrls) Console.WriteLine($"   - {u}");

                    string baseUrl = null;

                    foreach (var url in candidateUrls)
                    {
                        try
                        {
                            Console.WriteLine($"🔍 Thử URL: {url}");
                            driver.Navigate().GoToUrl(url);

                            wait.Until(d =>
                            {
                                try
                                {
                                    return ((IJavaScriptExecutor)d)
                                        .ExecuteScript("return document.readyState")
                                        .ToString() == "complete";
                                }
                                catch { return false; }
                            });

                            await Task.Delay(1500);

                            // Kiểm tra có form tra cứu không
                            var fkeyEls = driver.FindElements(By.Id("strFkey"));
                            if (fkeyEls.Count > 0 && fkeyEls[0].Displayed)
                            {
                                baseUrl = url;
                                Console.WriteLine($"✅ URL hợp lệ: {baseUrl}");
                                break;
                            }
                            else
                            {
                                Console.WriteLine($"⚠️ URL không có form tra cứu");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"⚠️ URL lỗi: {ex.Message}");
                        }
                    }

                    if (baseUrl == null)
                    {
                        Console.WriteLine("❌ Không tìm được URL VNPT hợp lệ");
                        XtraMessageBox.Show(
                            $"Không tìm thấy trang tra cứu VNPT cho MST {mst}",
                            "VNPT",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    // ═══════════════════════════════════════════════════════
                    // 🔁 VÒNG LẶP RETRY CAPTCHA — TỐI ĐA 5 LẦN
                    // ═══════════════════════════════════════════════════════
                    bool searchSuccess = false;
                    int maxCaptchaRetries = 5;

                    for (int attempt = 1; attempt <= maxCaptchaRetries && !searchSuccess; attempt++)
                    {
                        Console.WriteLine($"🔵 VNPT Captcha lần {attempt}/{maxCaptchaRetries}");

                        try
                        {
                            // ═══ 1. MỞ TRANG ═══
                            driver.Navigate().GoToUrl(baseUrl);

                            wait.Until(d =>
                            {
                                try
                                {
                                    return ((IJavaScriptExecutor)d)
                                        .ExecuteScript("return document.readyState")
                                        .ToString() == "complete";
                                }
                                catch { return false; }
                            });

                            await Task.Delay(1000);

                            // ═══ 2. NHẬP MÃ TRA CỨU (MÃ NHẬN HÓA ĐƠN) ═══
                            IWebElement fkey = wait.Until(d =>
                            {
                                try
                                {
                                    IWebElement el = d.FindElement(By.Id("strFkey"));
                                    return (el != null && el.Displayed && el.Enabled) ? el : null;
                                }
                                catch { return null; }
                            });

                            fkey.Clear();
                            fkey.SendKeys(mabm);
                            await Task.Delay(300);

                            // ═══════════════════════════════════════════════════
                            // ═══ 3. LẤY CAPTCHA BẰNG SCREENSHOT ELEMENT ═══
                            // ═══════════════════════════════════════════════════
                            byte[] captchaBytes = null;

                            // Tìm element captcha
                            IWebElement captchaImgEl = null;
                            string[] selectors = {
                        "img#captcha_img",
                        "img#captcha",
                        "img[src*='Captcha/Show']",
                        "img[src*='captcha']",
                        "div#captcha img",
                        "img[alt='captcha']"
                    };

                            foreach (var sel in selectors)
                            {
                                try
                                {
                                    var els = driver.FindElements(By.CssSelector(sel));
                                    foreach (var el in els)
                                    {
                                        if (el != null && el.Displayed && el.Size.Width > 30)
                                        {
                                            captchaImgEl = el;
                                            Console.WriteLine($"✅ Tìm captcha: {sel} ({el.Size.Width}x{el.Size.Height})");
                                            break;
                                        }
                                    }
                                    if (captchaImgEl != null) break;
                                }
                                catch { }
                            }

                            if (captchaImgEl == null)
                            {
                                Console.WriteLine("⚠️ Không tìm thấy element captcha — thử lại");
                                await Task.Delay(1000);
                                continue;
                            }

                            // Chụp trực tiếp element (chính xác 100% ảnh đang hiển thị)
                            captchaBytes = ScreenshotElement(driver, captchaImgEl);

                            // Lưu debug
                            if (captchaBytes != null && captchaBytes.Length > 100)
                            {
                                string debugPath = Path.Combine(Path.GetTempPath(), $"vnpt_captcha_{attempt}.png");
                                try { File.WriteAllBytes(debugPath, captchaBytes); } catch { }
                                Console.WriteLine($"📷 Captcha debug: {debugPath} ({captchaBytes.Length} bytes)");
                            }

                            if (captchaBytes == null || captchaBytes.Length < 100)
                            {
                                Console.WriteLine("⚠️ Không chụp được captcha — thử lại");
                                await Task.Delay(1000);
                                continue;
                            }

                            // ═══ 4. OCR CAPTCHA ═══
                            string code = Kingsofthp.OcrCaptcha(captchaBytes);

                            if (string.IsNullOrWhiteSpace(code) || code.Trim().Length != 4)
                            {
                                Console.WriteLine($"⚠️ OCR fail (code='{code}') — thử lại");
                                await Task.Delay(500);
                                continue;
                            }

                            code = code.Trim();
                            Console.WriteLine($"🤖 VNPT OCR → '{code}'");

                            // ═══ 5. NHẬP CAPTCHA ═══
                            IWebElement captchaInput = wait.Until(d =>
                            {
                                try
                                {
                                    IWebElement el = d.FindElement(By.Id("captch"));
                                    return (el != null && el.Displayed && el.Enabled) ? el : null;
                                }
                                catch { return null; }
                            });

                            captchaInput.Clear();
                            captchaInput.SendKeys(code);
                            await Task.Delay(300);

                            // ═══ 6. CLICK TÌM KIẾM ═══
                            IWebElement btnTimKiem = wait.Until(d =>
                            {
                                try
                                {
                                    IWebElement el = d.FindElement(
                                        By.CssSelector("button[type='submit'][name='submit']"));
                                    return (el != null && el.Displayed && el.Enabled) ? el : null;
                                }
                                catch { return null; }
                            });

                            ((IJavaScriptExecutor)driver).ExecuteScript(
                                "arguments[0].scrollIntoView({block: 'center'});", btnTimKiem);
                            await Task.Delay(200);
                            ((IJavaScriptExecutor)driver).ExecuteScript(
                                "arguments[0].click();", btnTimKiem);

                            Console.WriteLine("✅ Đã click Tìm kiếm");

                            // ═══ 7. KIỂM TRA URL ĐỔI HOẶC LỖI CAPTCHA ═══
                            bool urlChanged = false;
                            bool captchaError = false;
                            var urlDeadline = DateTime.Now.AddSeconds(12);

                            while (DateTime.Now < urlDeadline)
                            {
                                try
                                {
                                    if (driver.Url.Contains("/HomeNoLogin/SearchByFkey"))
                                    {
                                        urlChanged = true;
                                        break;
                                    }

                                    var errorMsgs = driver.FindElements(By.CssSelector(
                                        ".error, .alert, .text-danger, .validation-summary-errors, #captcha-error"));
                                    foreach (var err in errorMsgs)
                                    {
                                        if (!err.Displayed) continue;
                                        string txt = (err.Text ?? "").ToLower();
                                        if (txt.Contains("captcha") || txt.Contains("mã xác nhận") ||
                                            txt.Contains("không đúng") || txt.Contains("sai") ||
                                            txt.Contains("không chính xác"))
                                        {
                                            captchaError = true;
                                            Console.WriteLine($"❌ Server báo lỗi: '{err.Text}'");
                                            break;
                                        }
                                    }

                                    if (captchaError) break;
                                }
                                catch { }

                                await Task.Delay(300);
                            }

                            if (captchaError)
                            {
                                Console.WriteLine($"⚠️ Captcha '{code}' sai (server báo) — thử lại");
                                await Task.Delay(500);
                                continue;
                            }

                            if (!urlChanged)
                            {
                                Console.WriteLine($"⚠️ URL không đổi sau 12s — captcha '{code}' có thể sai, thử lại...");
                                await Task.Delay(1000);
                                continue;
                            }

                            Console.WriteLine("✅ Captcha đúng — đã vào trang kết quả");
                            searchSuccess = true;

                            // ═══ 8. TÌM LINK TẢI PDF ═══
                            IWebElement pdfLink = wait.Until(d =>
                            {
                                try
                                {
                                    IWebElement el = d.FindElement(
                                        By.CssSelector("a[title='Tải file pdf']"));
                                    return (el != null && el.Displayed && el.Enabled) ? el : null;
                                }
                                catch { return null; }
                            });

                            // ═══ 9. CLICK TẢI PDF ═══
                            var beforeFiles = new HashSet<string>(
                                Directory.GetFiles(downloadPath, "*.*")
                                         .Select(f => Path.GetFileName(f))
                            );

                            ((IJavaScriptExecutor)driver).ExecuteScript(
                                "arguments[0].scrollIntoView({block: 'center'});", pdfLink);
                            await Task.Delay(200);
                            ((IJavaScriptExecutor)driver).ExecuteScript(
                                "arguments[0].click();", pdfLink);

                            Console.WriteLine("✅ Đã click tải PDF, chờ file...");

                            // ═══ 10. CHỜ FILE PDF ═══
                            string pdfFile = WaitForPdf(shdon, mst, downloadPath, 20);

                            // Nếu lần 1 fail → click lại
                            if (pdfFile == null)
                            {
                                Console.WriteLine("⚠️ Lần 1 không có file — thử click lại...");
                                await Task.Delay(1000);

                                try
                                {
                                    ((IJavaScriptExecutor)driver).ExecuteScript(
                                        "document.querySelector(\"a[title='Tải file pdf']\")?.click();");
                                }
                                catch { }

                                pdfFile = WaitForPdf(shdon, mst, downloadPath, 20);
                            }

                            if (pdfFile == null)
                            {
                                Console.WriteLine("❌ Không tải được PDF sau 2 lần thử");
                                XtraMessageBox.Show(
                                    "Không thể tải PDF.", "VNPT",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                                return;
                            }

                            await Task.Delay(500);

                            // ═══ 11. ĐỔI TÊN ═══
                            string fileMoi = Path.Combine(
                                downloadPath,
                                $"{mst}_{shdon}_{KHHD}.pdf");

                            if (File.Exists(fileMoi))
                                File.Delete(fileMoi);

                            File.Move(pdfFile, fileMoi);

                            // ═══ 12. MỞ FILE ═══
                            if (toggle)
                            {
                                try
                                {
                                    Process.Start(new ProcessStartInfo
                                    {
                                        FileName = fileMoi,
                                        UseShellExecute = true
                                    });
                                }
                                catch { }
                            }

                            Console.WriteLine($"✅ Đã tải PDF VNPT: {fileMoi}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Lỗi captcha lần {attempt}: {ex.Message}");
                        }

                        if (!searchSuccess)
                        {
                            await Task.Delay(1000);
                        }
                    }

                    if (!searchSuccess)
                    {
                        Console.WriteLine("❌ Thử captcha nhiều lần không thành công");
                        XtraMessageBox.Show(
                            "Không thể vượt captcha VNPT sau nhiều lần thử.\r\n" +
                            "Vui lòng thử lại sau.",
                            "VNPT",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Lỗi tổng: {ex.Message}");
                    XtraMessageBox.Show(
                        "Lỗi tải hóa đơn VNPT:\r\n\r\n" + ex.Message,
                        "VNPT",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                finally
                {
                    await Task.Delay(1000);
                    try { driver?.Quit(); } catch { }
                    try { driver?.Dispose(); } catch { }
                    Console.WriteLine("Đã đóng trình duyệt.");
                }
            });
        }


        public async Task TaiHDVNPTPETROLIMEX(
     bool toggle,
     string mst,
     string shdon,
     string KHHD,
     string mabm,
     string savedPath,
     int type)
        {
            string downloadPath = "";

            if (type == 1)
            {
                downloadPath = Path.Combine(
                    savedPath,
                    $"HD{dateEdit1.DateTime.Year}",
                    "HDVao",
                    dateEdit1.DateTime.Month.ToString());
            }
            else if (type == 2)
            {
                downloadPath = Path.Combine(
                    savedPath,
                    $"HD{dateEdit1.DateTime.Year}",
                    "HDRa",
                    dateEdit1.DateTime.Month.ToString());
            }

            if (string.IsNullOrEmpty(downloadPath))
            {
                XtraMessageBox.Show("Chưa chọn đầu vào hoặc đầu ra.");
                return;
            }

            Directory.CreateDirectory(downloadPath);

            await Task.Run(async () =>
            {
                // ═══════════════════════════════════════════════════════
                // CHROME OPTIONS — CHỐNG BOT + CHỐNG POPUP DOWNLOAD
                // ═══════════════════════════════════════════════════════
                ChromeOptions options = new ChromeOptions();

               // options.AddArgument("--headless=new");
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

                options.AddUserProfilePreference("download.default_directory", downloadPath);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("safebrowsing.enabled", true);
                options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
                options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                options.AddUserProfilePreference("profile.default_content_setting_values.popups", 1);

                options.AddUserProfilePreference("profile.default_content_setting_values.media_stream", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.notifications", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.clipboard", 2);
                options.AddUserProfilePreference("credentials_enable_service", false);
                options.AddUserProfilePreference("profile.password_manager_enabled", false);
                options.AddUserProfilePreference("profile.password_manager_leak_detection", false);

                IWebDriver driver = null;
                try
                {
                    string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                    service.HideCommandPromptWindow = true;

                    driver = new ChromeDriver(service, options);
                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);

                    // ═══════════════════════════════════════════════════════
                    // 🛡️ CDP: DOWNLOAD BEHAVIOR + BYPASS CSP
                    // ═══════════════════════════════════════════════════════
                    try
                    {
                        if (driver is ChromeDriver chromeDriver)
                        {
                            chromeDriver.ExecuteCdpCommand(
                                "Page.setDownloadBehavior",
                                new Dictionary<string, object>
                                {
                            { "behavior", "allow" },
                            { "downloadPath", downloadPath }
                                });

                            chromeDriver.ExecuteCdpCommand("Page.enable", new Dictionary<string, object>());
                            chromeDriver.ExecuteCdpCommand(
                                "Page.setBypassCSP",
                                new Dictionary<string, object> { { "enabled", true } });

                            Console.WriteLine("✅ CDP download behavior đã bật");
                        }
                    }
                    catch (Exception cdpEx)
                    {
                        Console.WriteLine($"⚠️ CDP setup: {cdpEx.Message}");
                    }

                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

                    // ═══════════════════════════════════════════════════════
                    // 🌐 URL CỔNG PETROLIMEX
                    // ═══════════════════════════════════════════════════════
                    string baseUrl = "https://hoadon.petrolimex.com.vn/";
                    Console.WriteLine($"🌐 URL Petrolimex: {baseUrl}");

                    // ═══════════════════════════════════════════════════════
                    // 🔁 VÒNG LẶP RETRY CAPTCHA — TỐI ĐA 5 LẦN
                    // ═══════════════════════════════════════════════════════
                    bool searchSuccess = false;
                    int maxCaptchaRetries = 5;

                    for (int attempt = 1; attempt <= maxCaptchaRetries && !searchSuccess; attempt++)
                    {
                        Console.WriteLine($"🔵 Petrolimex Captcha lần {attempt}/{maxCaptchaRetries}");

                        try
                        {
                            // ═══ 1. MỞ TRANG ═══
                            driver.Navigate().GoToUrl(baseUrl);

                            wait.Until(d =>
                            {
                                try
                                {
                                    return ((IJavaScriptExecutor)d)
                                        .ExecuteScript("return document.readyState")
                                        .ToString() == "complete";
                                }
                                catch { return false; }
                            });

                            await Task.Delay(1000);

                            // ═══ 2. NHẬP MÃ TRA CỨU (FKEY) ═══
                            IWebElement fkey = wait.Until(d =>
                            {
                                try
                                {
                                    IWebElement el = d.FindElement(By.Id("strFkey"));
                                    return (el != null && el.Displayed && el.Enabled) ? el : null;
                                }
                                catch { return null; }
                            });

                            fkey.Clear();
                            fkey.SendKeys(mabm);
                            await Task.Delay(300);

                            // ═══════════════════════════════════════════════════
                            // ═══ 3. LẤY CAPTCHA BẰNG SCREENSHOT ELEMENT ═══
                            // ═══════════════════════════════════════════════════
                            byte[] captchaBytes = null;

                            // Tìm element captcha (Petrolimex dùng img.captcha_img)
                            IWebElement captchaImgEl = null;
                            string[] selectors = {
                        "img.captcha_img",
                        "img[src*='Captcha/Show']",
                        "img#captcha_img",
                        "img#captcha",
                        "img[src*='captcha']",
                        "div#captcha img",
                        "img[alt='captcha']"
                    };

                            foreach (var sel in selectors)
                            {
                                try
                                {
                                    var els = driver.FindElements(By.CssSelector(sel));
                                    foreach (var el in els)
                                    {
                                        if (el != null && el.Displayed && el.Size.Width > 30)
                                        {
                                            captchaImgEl = el;
                                            Console.WriteLine($"✅ Tìm captcha: {sel} ({el.Size.Width}x{el.Size.Height})");
                                            break;
                                        }
                                    }
                                    if (captchaImgEl != null) break;
                                }
                                catch { }
                            }

                            if (captchaImgEl == null)
                            {
                                Console.WriteLine("⚠️ Không tìm thấy element captcha — thử lại");
                                await Task.Delay(1000);
                                continue;
                            }

                            // Chụp trực tiếp element
                            captchaBytes = ScreenshotElement(driver, captchaImgEl);

                            // Lưu debug
                            if (captchaBytes != null && captchaBytes.Length > 100)
                            {
                                string debugPath = Path.Combine(Path.GetTempPath(), $"petro_captcha_{attempt}.png");
                                try { File.WriteAllBytes(debugPath, captchaBytes); } catch { }
                                Console.WriteLine($"📷 Captcha debug: {debugPath} ({captchaBytes.Length} bytes)");
                            }

                            if (captchaBytes == null || captchaBytes.Length < 100)
                            {
                                Console.WriteLine("⚠️ Không chụp được captcha — thử lại");
                                await Task.Delay(1000);
                                continue;
                            }

                            // ═══ 4. OCR CAPTCHA ═══
                            string code = Kingsofthp.OcrCaptcha(captchaBytes);

                            if (string.IsNullOrWhiteSpace(code) || code.Trim().Length != 4)
                            {
                                Console.WriteLine($"⚠️ OCR fail (code='{code}') — thử lại");
                                await Task.Delay(500);
                                continue;
                            }

                            code = code.Trim();
                            Console.WriteLine($"🤖 Petrolimex OCR → '{code}'");

                            // ═══ 5. NHẬP CAPTCHA ═══
                            IWebElement captchaInput = wait.Until(d =>
                            {
                                try
                                {
                                    IWebElement el = d.FindElement(By.Id("captch"));
                                    return (el != null && el.Displayed && el.Enabled) ? el : null;
                                }
                                catch { return null; }
                            });

                            captchaInput.Clear();
                            captchaInput.SendKeys(code);
                            await Task.Delay(300);

                            // ═══ 6. CLICK TÌM KIẾM ═══
                            // Petrolimex dùng input[type=submit][name=submit]
                            IWebElement btnTimKiem = wait.Until(d =>
                            {
                                try
                                {
                                    IWebElement el = d.FindElement(
                                        By.CssSelector("input[type='submit'][name='submit']"));
                                    return (el != null && el.Displayed && el.Enabled) ? el : null;
                                }
                                catch { return null; }
                            });

                            ((IJavaScriptExecutor)driver).ExecuteScript(
                                "arguments[0].scrollIntoView({block: 'center'});", btnTimKiem);
                            await Task.Delay(200);
                            ((IJavaScriptExecutor)driver).ExecuteScript(
                                "arguments[0].click();", btnTimKiem);

                            Console.WriteLine("✅ Đã click Tìm kiếm");

                            // ═══ 7. KIỂM TRA TABLE KẾT QUẢ XUẤT HIỆN ═══
                            bool tableFound = false;
                            bool captchaError = false;
                            var urlDeadline = DateTime.Now.AddSeconds(15);

                            while (DateTime.Now < urlDeadline)
                            {
                                try
                                {
                                    // Kiểm tra table kết quả có link tải PDF không
                                    var pdfLinks = driver.FindElements(
                                        By.CssSelector("a[onclick*='ajxCall4PortalPDF']"));
                                    if (pdfLinks.Count > 0 && pdfLinks[0].Displayed)
                                    {
                                        tableFound = true;
                                        break;
                                    }

                                    // Kiểm tra lỗi captcha
                                    var errorMsgs = driver.FindElements(By.CssSelector(
                                        ".error, .alert, .text-danger, .validation-summary-errors, #CaptchaError"));
                                    foreach (var err in errorMsgs)
                                    {
                                        if (!err.Displayed) continue;
                                        string txt = (err.Text ?? "").ToLower();
                                        if (txt.Contains("captcha") || txt.Contains("mã xác nhận") ||
                                            txt.Contains("không đúng") || txt.Contains("sai") ||
                                            txt.Contains("không chính xác"))
                                        {
                                            captchaError = true;
                                            Console.WriteLine($"❌ Server báo lỗi: '{err.Text}'");
                                            break;
                                        }
                                    }

                                    if (captchaError) break;
                                }
                                catch { }

                                await Task.Delay(300);
                            }

                            if (captchaError)
                            {
                                Console.WriteLine($"⚠️ Captcha '{code}' sai (server báo) — thử lại");
                                await Task.Delay(500);
                                continue;
                            }

                            if (!tableFound)
                            {
                                Console.WriteLine($"⚠️ Không thấy table kết quả sau 15s — captcha '{code}' có thể sai, thử lại...");
                                await Task.Delay(1000);
                                continue;
                            }

                            Console.WriteLine("✅ Captcha đúng — table kết quả đã xuất hiện");
                            searchSuccess = true;

                            // ═══ 8. TÌM LINK TẢI PDF ═══
                            IWebElement pdfLink = wait.Until(d =>
                            {
                                try
                                {
                                    IWebElement el = d.FindElement(
                                        By.CssSelector("a[onclick*='ajxCall4PortalPDF']"));
                                    return (el != null && el.Displayed) ? el : null;
                                }
                                catch { return null; }
                            });

                            // ═══ 9. CLICK TẢI PDF ═══
                            var beforeFiles = new HashSet<string>(
                                Directory.GetFiles(downloadPath, "*.*")
                                         .Select(f => Path.GetFileName(f))
                            );

                            ((IJavaScriptExecutor)driver).ExecuteScript(
                                "arguments[0].scrollIntoView({block: 'center'});", pdfLink);
                            await Task.Delay(200);
                            ((IJavaScriptExecutor)driver).ExecuteScript(
                                "arguments[0].click();", pdfLink);

                            Console.WriteLine("✅ Đã click tải PDF, chờ file...");

                            // ═══ 10. CHỜ FILE PDF ═══
                            string pdfFile = WaitForPdf(shdon, mst, downloadPath, 20);

                            // Nếu lần 1 fail → click lại
                            if (pdfFile == null)
                            {
                                Console.WriteLine("⚠️ Lần 1 không có file — thử click lại...");
                                await Task.Delay(1000);

                                try
                                {
                                    ((IJavaScriptExecutor)driver).ExecuteScript(
                                        "document.querySelector(\"a[onclick*='ajxCall4PortalPDF']\")?.click();");
                                }
                                catch { }

                                pdfFile = WaitForPdf(shdon, mst, downloadPath, 20);
                            }

                            if (pdfFile == null)
                            {
                                Console.WriteLine("❌ Không tải được PDF sau 2 lần thử");
                                XtraMessageBox.Show(
                                    "Không thể tải PDF.", "Petrolimex",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                                return;
                            }

                            await Task.Delay(500);

                            // ═══ 11. ĐỔI TÊN ═══
                            string fileMoi = Path.Combine(
                                downloadPath,
                                $"{mst}_{shdon}_{KHHD}.pdf");

                            if (File.Exists(fileMoi))
                                File.Delete(fileMoi);

                            File.Move(pdfFile, fileMoi);

                            // ═══ 12. MỞ FILE ═══
                            if (toggle)
                            {
                                try
                                {
                                    Process.Start(new ProcessStartInfo
                                    {
                                        FileName = fileMoi,
                                        UseShellExecute = true
                                    });
                                }
                                catch { }
                            }

                            Console.WriteLine($"✅ Đã tải PDF Petrolimex: {fileMoi}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Lỗi captcha lần {attempt}: {ex.Message}");
                        }

                        if (!searchSuccess)
                        {
                            await Task.Delay(1000);
                        }
                    }

                    if (!searchSuccess)
                    {
                        Console.WriteLine("❌ Thử captcha nhiều lần không thành công");
                        XtraMessageBox.Show(
                            "Không thể vượt captcha Petrolimex sau nhiều lần thử.\r\n" +
                            "Vui lòng thử lại sau.",
                            "Petrolimex",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Lỗi tổng: {ex.Message}");
                    XtraMessageBox.Show(
                        "Lỗi tải hóa đơn Petrolimex:\r\n\r\n" + ex.Message,
                        "Petrolimex",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                finally
                {
                    await Task.Delay(1000);
                    try { driver?.Quit(); } catch { }
                    try { driver?.Dispose(); } catch { }
                    Console.WriteLine("Đã đóng trình duyệt.");
                }
            });
        }
        private byte[] ScreenshotElement(IWebDriver driver, IWebElement element)
        {
            try
            {
                // Cuộn element vào giữa màn hình
                ((IJavaScriptExecutor)driver).ExecuteScript(
                    "arguments[0].scrollIntoView({block: 'center'});", element);
                System.Threading.Thread.Sleep(200);

                // Chụp toàn trang
                var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                using (var ms = new MemoryStream(screenshot.AsByteArray))
                {
                    var fullImage = System.Drawing.Image.FromStream(ms);

                    // Lấy vị trí + kích thước element
                    var location = element.Location;
                    var size = element.Size;

                    // Cắt vùng element
                    var rect = new System.Drawing.Rectangle(
                        location.X, location.Y, size.Width, size.Height);

                    using (var bmp = new System.Drawing.Bitmap(rect.Width, rect.Height))
                    {
                        using (var g = System.Drawing.Graphics.FromImage(bmp))
                        {
                            g.DrawImage(fullImage,
                                new System.Drawing.Rectangle(0, 0, rect.Width, rect.Height),
                                rect,
                                System.Drawing.GraphicsUnit.Pixel);
                        }

                        // Chuyển thành byte[]
                        using (var outMs = new MemoryStream())
                        {
                            bmp.Save(outMs, System.Drawing.Imaging.ImageFormat.Png);
                            return outMs.ToArray();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ ScreenshotElement fail: {ex.Message}");
                return null;
            }
        }
        public async Task TaiMINVOICE(bool toggle, string mst, string shdon, string KHHDon, string mabimat, string savedPath, int type)
        {
            string url = "https://tracuuhoadon.minvoice.com.vn/tra-cuu-hoa-don";
            string downloadPath = "";

            await Task.Run(async () =>
            {
                if (type==1)
                    downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDVao\\" + dateEdit1.DateTime.Month;
                if (type==2)
                    downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDRa\\" + dateEdit1.DateTime.Month;

                Directory.CreateDirectory(downloadPath);

                ChromeOptions options = new ChromeOptions();
                options.AddArgument("--headless=new");
                options.AddArgument("--disable-gpu");
                options.AddArgument("--window-size=1920,1080");
                options.AddArgument("--disable-notifications");
                options.AddArgument("--disable-infobars");
                options.AddArgument("--no-first-run");
                options.AddArgument("--no-default-browser-check");
                options.AddArgument("--disable-features=Translate,BackForwardCache,OptimizationHints");
                options.AddArgument("--disable-blink-features=AutomationControlled");

                // === CHẶN QUẢNG CÁO BẰNG EXTENSION uBlock Origin ===
                // Đường dẫn tới thư mục extension đã giải nén (chứa manifest.json)
                // Tải uBlock Origin: https://github.com/gorhill/uBlock/releases
                // Giải nén file .crx (đổi đuôi thành .zip rồi unzip) hoặc dùng Chrome Web Store unpacked
                string ublockPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Extensions", "uBlock");
                if (Directory.Exists(ublockPath))
                {
                    options.AddArgument("--load-extension=" + ublockPath);
                }

                // Download preferences
                options.AddUserProfilePreference("download.default_directory", downloadPath);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("safebrowsing.enabled", true);
                options.AddUserProfilePreference("profile.default_content_setting_values.popups", 2); // chặn popup

                IWebDriver driver = null;
                try
                {
                    string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                    service.HideCommandPromptWindow = true; // Ẩn cửa sổ cmd đen

                    driver = new ChromeDriver(service, options);

                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                    driver.Navigate().GoToUrl(url);
                     
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

                    // ============================================================
                    // 1. Điền form tra cứu
                    // ============================================================
                    var taxCodeInput = wait.Until(d => d.FindElement(By.Id("taxCode")));
                    taxCodeInput.Clear();
                    taxCodeInput.SendKeys(mst);

                    var checkCodeInput = wait.Until(d => d.FindElement(By.Id("checkCode")));
                    checkCodeInput.Clear();
                    checkCodeInput.SendKeys(mabimat);

                    // Lưu window chính trước khi click
                    string mainWindow = driver.CurrentWindowHandle;

                    // ============================================================
                    // 2. Nhấn nút Tra cứu — dùng JS click tránh overlay
                    // ============================================================
                    var btnSearch = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));

                    // Xoá overlay quảng cáo nếu có
                    try
                    {
                        js.ExecuteScript(@"
                document.querySelectorAll(
                    'iframe[id*=google_ads], ins.adsbygoogle, div[class*=popup], div[id*=popup], div[class*=overlay]'
                ).forEach(function(e){ e.remove(); });
            ");
                    }
                    catch { }

                    js.ExecuteScript("arguments[0].click();", btnSearch);

                    // ============================================================
                    // 3. Dọn tab quảng cáo nếu bị mở (Shopee, ...)
                    // ============================================================
                    Thread.Sleep(2000);
                    var handles = driver.WindowHandles;
                    if (handles.Count > 1)
                    {
                        foreach (var h in handles)
                        {
                            if (h != mainWindow)
                            {
                                try
                                {
                                    driver.SwitchTo().Window(h);
                                    string advUrl = driver.Url;
                                    Console.WriteLine("🚫 Đóng tab quảng cáo: " + advUrl);
                                    driver.Close();
                                }
                                catch { }
                            }
                        }
                        driver.SwitchTo().Window(mainWindow);
                    }

                    // Nếu URL hiện tại bị đổi sang domain khác → quay lại
                    if (!driver.Url.Contains("minvoice.com.vn"))
                    {
                        Console.WriteLine("⚠ Bị chuyển hướng sang: " + driver.Url + " → quay lại");
                        driver.Navigate().Back();
                        Thread.Sleep(1000);
                    }

                    // ============================================================
                    // 4. Chờ modal "Xem hóa đơn (PDF)" xuất hiện
                    // ============================================================
                    var modal = wait.Until(d =>
                    {
                        try
                        {
                            var m = d.FindElement(By.CssSelector("div.ant-modal[role='dialog']"));
                            if (!m.Displayed) return null;
                            var iframe = m.FindElements(By.CssSelector("iframe[title='Xem hóa đơn (PDF)']"));
                            if (iframe.Count == 0) return null;
                            return m;
                        }
                        catch (NoSuchElementException) { return null; }
                        catch (StaleElementReferenceException) { return null; }
                    });

                    // ============================================================
                    // 5. Chờ iframe PDF load (src bắt đầu bằng blob:)
                    // ============================================================
                    wait.Until(d =>
                    {
                        try
                        {
                            var iframe = modal.FindElement(By.CssSelector("iframe[title='Xem hóa đơn (PDF)']"));
                            var src = iframe.GetAttribute("src");
                            return !string.IsNullOrEmpty(src) && src.StartsWith("blob:");
                        }
                        catch (StaleElementReferenceException) { return false; }
                        catch (NoSuchElementException) { return false; }
                    });

                    // ============================================================
                    // 6. Tìm nút PDF (phân biệt với nút XML cùng icon download)
                    // ============================================================
                    IWebElement btnPdf = wait.Until(d =>
                    {
                        try
                        {
                            var buttons = modal.FindElements(By.CssSelector("button.ant-btn"));
                            foreach (var b in buttons)
                            {
                                try
                                {
                                    var dlIcon = b.FindElements(By.CssSelector("span.anticon-download"));
                                    if (dlIcon.Count == 0) continue;

                                    var spans = b.FindElements(By.TagName("span"));
                                    foreach (var s in spans)
                                    {
                                        if (s.Text.Trim() == "PDF")
                                            return b;
                                    }
                                }
                                catch (StaleElementReferenceException) { }
                            }
                            return null;
                        }
                        catch (StaleElementReferenceException) { return null; }
                    });

                    // ============================================================
                    // 7. Click tải PDF (JS click tránh overlay)
                    // ============================================================
                    js.ExecuteScript("arguments[0].click();", btnPdf);

                    // ============================================================
                    // 8. Chờ file PDF tải xong
                    // ============================================================
                    string pdfFile = WaitForPdf(shdon, mst, downloadPath, 10);
                    if (pdfFile == null)
                    {
                        XtraMessageBox.Show("Không thể tải");
                        driver.Quit();
                    }
                    Thread.Sleep(500);

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
                    Console.WriteLine("❌ Lỗi TaiMINVOICE: " + ex.Message);
                    throw;
                }
                finally
                {
                    // Luôn đóng driver kể cả khi lỗi
                    try { if (driver != null) driver.Quit(); } catch { }
                    try { if (driver != null) driver.Dispose(); } catch { }
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

                }
            });
        }
        public async Task TaiBKAV(bool toggle, string mst, string shdon, string KHHDon, string mabimat, string savedPath, int type)
        {
            string url = $"https://van.ehoadon.vn/Lookup?InvoiceGUID={mabimat}";
            string downloadPath = "";

            await Task.Run(async () =>
            {
                // ═══ 1. XÁC ĐỊNH THƯ MỤC ═══
                if (type == 1)
                    downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDVao\\" + dateEdit1.DateTime.Month;
                else if (type == 2)
                    downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDRa\\" + dateEdit1.DateTime.Month;
                else
                    downloadPath = savedPath + $"\\HD{dateEdit1.DateTime.Year}" + "\\HDKhac\\" + dateEdit1.DateTime.Month;

                Directory.CreateDirectory(downloadPath);

                ChromeOptions options = new ChromeOptions();
                options.AddArgument("--headless=new");
                options.AddArgument("--disable-blink-features=AutomationControlled");
                options.AddExcludedArgument("enable-automation");
                options.AddAdditionalOption("useAutomationExtension", false);
                options.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36");
                options.AddArgument($"--download-default-directory={downloadPath}");

                options.AddArgument("--disable-features=DownloadBubble,DownloadBubbleV2," +
                                    "OptimizationHints,MediaRouter,PrivacySandboxSettings4," +
                                    "PrivacySandboxAdsAPIs,InterestFeedContentSuggestions,Translate," +
                                    "AcceptCHFrame,ChromeLabs,CalculateNativeWinOcclusion," +
                                    "WebAuthentication,WebBluetooth,WebUsb,WebSerial,WebHID,WebNFC," +
                                    "FileSystemAccessAPI,IdleDetection,SyncDisabled");
                options.AddArgument("--disable-notifications");
                options.AddArgument("--disable-infobars");
                options.AddArgument("--disable-popup-blocking");
                options.AddArgument("--no-first-run");
                options.AddArgument("--no-default-browser-check");
                options.AddArgument("--disable-default-apps");
                options.AddArgument("--disable-gpu");
                options.AddArgument("--no-sandbox");
                options.AddArgument("--disable-dev-shm-usage");
                options.AddArgument("--window-size=1920,1080");
                options.AddArgument("--disable-extensions");
                options.AddArgument("--disable-web-security");
                options.AddArgument("--allow-running-insecure-content");

                options.AddUserProfilePreference("download.default_directory", downloadPath);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("safebrowsing.enabled", false);
                options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
                options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                options.AddUserProfilePreference("profile.default_content_setting_values.popups", 1);
                options.AddUserProfilePreference("profile.default_content_setting_values.media_stream", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.notifications", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.clipboard", 2);
                options.AddUserProfilePreference("credentials_enable_service", false);
                options.AddUserProfilePreference("profile.password_manager_enabled", false);

                IWebDriver driver = null;
                try
                {
                    string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                    service.HideCommandPromptWindow = true;

                    driver = new ChromeDriver(service, options);
                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                    // ═══ CDP ═══
                    if (driver is ChromeDriver chromeDriver)
                    {
                        // Browser level
                        try
                        {
                            chromeDriver.ExecuteCdpCommand(
                                "Browser.setDownloadBehavior",
                                new Dictionary<string, object>
                                {
                            { "behavior", "allow" },
                            { "downloadPath", downloadPath },
                            { "eventsEnabled", true }
                                });
                            Console.WriteLine("✅ Browser.setDownloadBehavior OK");
                        }
                        catch { }

                        // Page level (override browser)
                        try
                        {
                            chromeDriver.ExecuteCdpCommand(
                                "Page.setDownloadBehavior",
                                new Dictionary<string, object>
                                {
                            { "behavior", "allow" },
                            { "downloadPath", downloadPath }
                                });
                            Console.WriteLine("✅ Page.setDownloadBehavior OK");
                        }
                        catch { }

                        chromeDriver.ExecuteCdpCommand("Page.enable", new Dictionary<string, object>());
                        chromeDriver.ExecuteCdpCommand("Page.setBypassCSP",
                            new Dictionary<string, object> { { "enabled", true } });
                    }

                    driver.Navigate().GoToUrl(url);

                    wait.Until(d =>
                    {
                        try
                        {
                            return ((IJavaScriptExecutor)d)
                                .ExecuteScript("return document.readyState")
                                .ToString() == "complete";
                        }
                        catch { return false; }
                    });

                    // ═══ TÌM NÚT DOWNLOAD ═══
                    IWebElement btnDownload = wait.Until(d =>
                    {
                        try
                        {
                            IWebElement el = d.FindElement(By.Id("btnDownload"));
                            return el.Displayed && el.Enabled ? el : null;
                        }
                        catch { return null; }
                    });

                    Actions actions = new Actions(driver);
                    actions.MoveToElement(btnDownload).Perform();
                    await Task.Delay(300);

                    wait.Until(d =>
                    {
                        try
                        {
                            IWebElement menu = d.FindElement(By.CssSelector("#divDownloads ul.dropdown-menu"));
                            return menu.Displayed;
                        }
                        catch { return false; }
                    });

                    IWebElement pdfButton = wait.Until(d =>
                    {
                        try
                        {
                            IWebElement el = d.FindElement(By.Id("LinkDownPDF"));
                            return el.Displayed && el.Enabled ? el : null;
                        }
                        catch { return null; }
                    });

                    // ═══ LẤY THÔNG TIN ═══
                    string pdfUrl = pdfButton.GetAttribute("href") ?? "";
                    string dataUrl = pdfButton.GetAttribute("data-url") ?? "";
                    string onclick = pdfButton.GetAttribute("onclick") ?? "";

                    Console.WriteLine($"📎 href={pdfUrl}");
                    Console.WriteLine($"📎 data-url={dataUrl}");
                    Console.WriteLine($"📎 onclick={onclick}");

                    bool downloaded = false;
                    string fileMoi = Path.Combine(downloadPath, $"{mst}_{shdon}_{KHHDon}.pdf");

                    // ═══════════════════════════════════════════════════════
                    // ✅ CÁCH 1: TRÍCH XUẤT URL TỪ onclick
                    // ═══════════════════════════════════════════════════════
                    string relativePath = "";

                    if (!string.IsNullOrEmpty(onclick))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(
                            onclick,
                            @"DownloadFile\s*\(\s*['""]([^'""]+)['""]");

                        if (match.Success)
                        {
                            relativePath = match.Groups[1].Value;
                            Console.WriteLine($"✅ Trích xuất path: {relativePath}");
                        }
                    }

                    // Thử tải từ nhiều base URL
                    if (!string.IsNullOrEmpty(relativePath))
                    {
                        // BKAV có thể dùng nhiều base URL khác nhau
                        string[] baseUrls = new[]
                        {
                    "https://van.ehoadon.vn/",
                    "https://van.ehoadon.vn/api/",
                    "https://van.ehoadon.vn/api/download/",
                    "https://van.ehoadon.vn/File/",
                    "https://van.ehoadon.vn/Download/",
                    "https://van.ehoadon.vn/Upload/",
                    "https://van.ehoadon.vn/Files/",
                };

                        // Lấy cookie
                        var cookies = driver.Manage().Cookies.AllCookies;
                        string cookieStr = string.Join("; ", cookies.Select(c => $"{c.Name}={c.Value}"));

                        foreach (var baseUrl in baseUrls)
                        {
                            if (downloaded) break;

                            string fullUrl = baseUrl + relativePath;
                            Console.WriteLine($"📥 Thử: {fullUrl}");

                            try
                            {
                                using (var httpClient = new HttpClient())
                                {
                                    httpClient.Timeout = TimeSpan.FromSeconds(15);
                                    httpClient.DefaultRequestHeaders.Add("User-Agent",
                                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
                                        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36");
                                    httpClient.DefaultRequestHeaders.Add("Referer", url);
                                    httpClient.DefaultRequestHeaders.Add("Accept", "application/pdf,*/*");
                                    if (!string.IsNullOrEmpty(cookieStr))
                                        httpClient.DefaultRequestHeaders.Add("Cookie", cookieStr);

                                    var response = await httpClient.GetAsync(fullUrl);

                                    if (response.IsSuccessStatusCode)
                                    {
                                        byte[] fileBytes = await response.Content.ReadAsByteArrayAsync();

                                        // Check magic bytes PDF: %PDF
                                        if (fileBytes.Length > 500 &&
                                            fileBytes[0] == 0x25 &&   // %
                                            fileBytes[1] == 0x50 &&   // P
                                            fileBytes[2] == 0x44 &&   // D
                                            fileBytes[3] == 0x46)     // F
                                        {
                                            if (File.Exists(fileMoi)) File.Delete(fileMoi);
                                            File.WriteAllBytes(fileMoi, fileBytes);

                                            Console.WriteLine($"✅ PDF OK: {fileMoi} ({fileBytes.Length} bytes)");
                                            downloaded = true;

                                            if (toggle)
                                            {
                                                try
                                                {
                                                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                                    {
                                                        FileName = fileMoi,
                                                        UseShellExecute = true
                                                    });
                                                }
                                                catch { }
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine($"⚠️ Không phải PDF ({fileBytes.Length} bytes)");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"⚠️ HTTP {(int)response.StatusCode}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"⚠️ Fail: {ex.Message}");
                            }
                        }
                    }

                    // ═══════════════════════════════════════════════════════
                    // ✅ CÁCH 2: NẾU CÁCH 1 FAIL → CLICK + TÌM FILE Ở CẢ 2 FOLDER
                    // ═══════════════════════════════════════════════════════
                    if (!downloaded)
                    {
                        Console.WriteLine("📥 Fallback: Click LinkDownPDF...");

                        string defaultDownloads = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                            "Downloads");

                        var beforeInDownload = new HashSet<string>(
                            Directory.Exists(downloadPath)
                                ? Directory.GetFiles(downloadPath).Select(Path.GetFileName)
                                : new string[0]);

                        var beforeInDefault = new HashSet<string>(
                            Directory.Exists(defaultDownloads)
                                ? Directory.GetFiles(defaultDownloads).Select(Path.GetFileName)
                                : new string[0]);

                        // Click
                        try
                        {
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", pdfButton);
                        }
                        catch
                        {
                            pdfButton.Click();
                        }

                        // Tìm file ở cả 2 folder
                        string pdfFile = WaitForPdfInBothFolders(
                            downloadPath, defaultDownloads,
                            beforeInDownload, beforeInDefault, 25);

                        if (pdfFile == null)
                        {
                            Console.WriteLine("⚠️ Lần 1 không có file — thử lại");
                            await Task.Delay(2000);

                            try
                            {
                                ((IJavaScriptExecutor)driver).ExecuteScript(
                                    "document.getElementById('LinkDownPDF').click();");
                            }
                            catch { }

                            pdfFile = WaitForPdfInBothFolders(
                                downloadPath, defaultDownloads,
                                beforeInDownload, beforeInDefault, 25);
                        }

                        if (pdfFile != null && File.Exists(pdfFile))
                        {
                            if (!string.Equals(pdfFile, fileMoi, StringComparison.OrdinalIgnoreCase))
                            {
                                if (File.Exists(fileMoi)) File.Delete(fileMoi);
                                try
                                {
                                    File.Move(pdfFile, fileMoi);
                                    Console.WriteLine($"✅ Đã lưu: {fileMoi}");
                                }
                                catch
                                {
                                    File.Copy(pdfFile, fileMoi, true);
                                }
                            }

                            if (toggle)
                            {
                                try
                                {
                                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                    {
                                        FileName = fileMoi,
                                        UseShellExecute = true
                                    });
                                }
                                catch { }
                            }

                            downloaded = true;
                        }
                        else
                        {
                            Console.WriteLine("❌ Không tìm thấy file");
                            XtraMessageBox.Show("Không thể tải hóa đơn BKAV");
                        }
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
                }
            });
        }

        // ═══════════════════════════════════════════════════════════════
        // 🔧 HELPER: TÌM FILE PDF MỚI Ở CẢ 2 FOLDER
        // ═══════════════════════════════════════════════════════════════
        private string WaitForPdfInBothFolders(
            string expectedFolder,
            string defaultFolder,
            HashSet<string> beforeInExpected,
            HashSet<string> beforeInDefault,
            int timeoutSeconds)
        {
            DateTime start = DateTime.Now;

            while ((DateTime.Now - start).TotalSeconds < timeoutSeconds)
            {
                try
                {
                    // ═══ CHECK 1: FOLDER ĐÍCH TRƯỚC ═══
                    if (Directory.Exists(expectedFolder))
                    {
                        // Còn file .crdownload → đang tải
                        var downloading = Directory.GetFiles(expectedFolder, "*.crdownload");
                        if (downloading.Length > 0)
                        {
                            Thread.Sleep(500);
                            continue;
                        }

                        var newPdf = Directory.GetFiles(expectedFolder, "*.pdf")
                            .Where(f => !beforeInExpected.Contains(Path.GetFileName(f)))
                            .OrderByDescending(File.GetCreationTime)
                            .FirstOrDefault();

                        if (newPdf != null && IsFileReady(newPdf))
                        {
                            Console.WriteLine($"✅ File ở downloadPath: {newPdf}");
                            return newPdf;
                        }
                    }

                    // ═══ CHECK 2: FOLDER DOWNLOADS MẶC ĐỊNH ═══
                    if (Directory.Exists(defaultFolder))
                    {
                        var downloading = Directory.GetFiles(defaultFolder, "*.crdownload");
                        if (downloading.Length > 0)
                        {
                            Thread.Sleep(500);
                            continue;
                        }

                        var newPdf = Directory.GetFiles(defaultFolder, "*.pdf")
                            .Where(f => !beforeInDefault.Contains(Path.GetFileName(f)))
                            .Where(f =>
                            {
                                // Chỉ lấy file mới trong 3 phút
                                var age = DateTime.Now - File.GetCreationTime(f);
                                return age.TotalMinutes < 3;
                            })
                            .OrderByDescending(File.GetCreationTime)
                            .FirstOrDefault();

                        if (newPdf != null && IsFileReady(newPdf))
                        {
                            Console.WriteLine($"⚠️ File ở Downloads mặc định: {newPdf}");
                            return newPdf;
                        }
                    }
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }

                Thread.Sleep(500);
            }

            Console.WriteLine($"⏱ Hết {timeoutSeconds}s không tìm thấy PDF");
            return null;
        }

        // ═══════════════════════════════════════════════════════════════
        // 🔧 HELPER: KIỂM TRA FILE ĐÃ GHI XONG CHƯA
        // ═══════════════════════════════════════════════════════════════
        private bool IsFileReady(string path)
        {
            try
            {
                // Check 1: có mở được không
                using (var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    long size1 = fs.Length;
                    fs.Close();

                    // Check 2: size ổn định
                    Thread.Sleep(300);
                    long size2 = new FileInfo(path).Length;

                    return size1 == size2 && size1 > 0;
                }
            }
            catch
            {
                return false;
            }
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
                options.AddArgument("--headless=new");
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
            // ===== 1. XÁC ĐỊNH THƯ MỤC ĐÍCH =====
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

                // Thư mục tạm
                string tempDownloadFolder = Path.Combine(Path.GetTempPath(), "MisaTemp", Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempDownloadFolder);

                // ═══════════════════════════════════════════════════════
                // CHROME OPTIONS — CHỐNG POPUP DOWNLOAD
                // ═══════════════════════════════════════════════════════
                ChromeOptions options = new ChromeOptions();

                //options.AddArgument("--headless");
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

                // Chặn popup + permission
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

                // Download
                options.AddUserProfilePreference("download.default_directory", tempDownloadFolder);
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
                options.AddUserProfilePreference("profile.default_content_setting_values.clipboard", 2);
                options.AddUserProfilePreference("credentials_enable_service", false);
                options.AddUserProfilePreference("profile.password_manager_enabled", false);
                options.AddUserProfilePreference("profile.password_manager_leak_detection", false);

                IWebDriver driver = null;
                try
                {
                    string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                    service.HideCommandPromptWindow = true;

                    driver = new ChromeDriver(service, options);
                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);

                    // ═══════════════════════════════════════════════════════
                    // 🛡️ CDP: DOWNLOAD BEHAVIOR + BYPASS CSP
                    // ═══════════════════════════════════════════════════════
                    try
                    {
                        if (driver is ChromeDriver chromeDriver)
                        {
                            chromeDriver.ExecuteCdpCommand(
                                "Page.setDownloadBehavior",
                                new Dictionary<string, object>
                                {
                            { "behavior", "allow" },
                            { "downloadPath", tempDownloadFolder }
                                });

                            chromeDriver.ExecuteCdpCommand("Page.enable", new Dictionary<string, object>());
                            chromeDriver.ExecuteCdpCommand(
                                "Page.setBypassCSP",
                                new Dictionary<string, object> { { "enabled", true } });

                            Console.WriteLine("✅ CDP download behavior đã bật");
                        }
                    }
                    catch (Exception cdpEx)
                    {
                        Console.WriteLine($"⚠️ CDP setup: {cdpEx.Message}");
                    }

                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                    // ===== 1. ĐIỀU HƯỚNG =====
                    Console.WriteLine("Đang truy cập trang tra cứu...");
                    driver.Navigate().GoToUrl("https://www.meinvoice.vn/tra-cuu/");
                    //await Task.Delay(200);

                    // ===== 2. NHẬP MÃ BÍ MẬT =====
                    IWebElement inputCode = wait.Until(drv =>
                    {
                        try
                        {
                            var el = drv.FindElement(By.Id("txtCode"));
                            return (el != null && el.Displayed) ? el : null;
                        }
                        catch { return null; }
                    });

                    inputCode.Clear();
                    inputCode.SendKeys(maBiMat);
                    Console.WriteLine($"Đã nhập mã: {maBiMat}");
                    //await Task.Delay(200);

                    // ═══════════════════════════════════════════════════════
                    // 3. XỬ LÝ CAPTCHA (nếu có)
                    // ═══════════════════════════════════════════════════════
                  //  await HandleMisaCaptcha(driver, wait);

                    // ===== 4. CLICK NÚT TÌM KIẾM =====
                    try
                    {
                        var searchBtn = driver.FindElement(By.Id("btnSearchInvoice"));
                        ((IJavaScriptExecutor)driver).ExecuteScript(
                            "arguments[0].scrollIntoView({block: 'center'});", searchBtn);
                        await Task.Delay(200);
                        ((IJavaScriptExecutor)driver).ExecuteScript(
                            "arguments[0].click();", searchBtn);
                        Console.WriteLine("✅ Đã click nút Tìm kiếm");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠️ Không click được nút tìm kiếm: {ex.Message}");
                        return;
                    }

                   // await Task.Delay(400);

                    // ===== 5. ĐỢI POPUP =====
                    IWebElement popupDiv = null;
                    try
                    {
                        popupDiv = wait.Until(drv =>
                        {
                            try
                            {
                                var el = drv.FindElement(By.Id("showContentPopupInvoice"));
                                return (el != null && el.Displayed) ? el : null;
                            }
                            catch { return null; }
                        });
                        Console.WriteLine("✅ Đã tìm thấy popup hóa đơn");
                    }
                    catch (WebDriverTimeoutException)
                    {
                        Console.WriteLine("⚠️ Không tìm thấy popup hóa đơn — có thể captcha sai hoặc mã không đúng");
                        return;
                    }

                    // ===== 6. CLICK "TẢI HÓA ĐƠN" =====
                    try
                    {
                        var downloadSpan = popupDiv.FindElement(By.CssSelector("span.download-invoice"));
                        ((IJavaScriptExecutor)driver).ExecuteScript(
                            "arguments[0].click();", downloadSpan);
                        Console.WriteLine("✅ Đã click 'Tải hóa đơn'");
                    }
                    catch (NoSuchElementException)
                    {
                        Console.WriteLine("⚠️ Không tìm thấy nút 'Tải hóa đơn'");
                    }

                  //  await Task.Delay(800);

                    // ===== 8. CHỤP DANH SÁCH FILE TRƯỚC =====
                    var beforeFiles = new HashSet<string>(
                        Directory.GetFiles(tempDownloadFolder, "*.*")
                                 .Select(f => Path.GetFileName(f))
                    );

                    // ===== 9. CLICK "TẢI HÓA ĐƠN DẠNG PDF" =====
                    bool pdfClicked = false;
                    try
                    {
                        var pdfOption = wait.Until(drv =>
                        {
                            try
                            {
                                var el = drv.FindElement(By.CssSelector("div.dm-item.pdf.txt-download-pdf"));
                                return (el != null && el.Displayed) ? el : null;
                            }
                            catch { return null; }
                        });
                        ((IJavaScriptExecutor)driver).ExecuteScript(
                            "arguments[0].click();", pdfOption);
                        Console.WriteLine("✅ Đã click 'Tải hóa đơn dạng PDF'");
                        pdfClicked = true;
                    }
                    catch (WebDriverTimeoutException)
                    {
                        try
                        {
                            var optionsList = driver.FindElements(By.CssSelector("div.dm-item"));
                            foreach (var opt in optionsList)
                            {
                                if (opt.Text.Contains("PDF") || opt.Text.Contains("pdf"))
                                {
                                    ((IJavaScriptExecutor)driver).ExecuteScript(
                                        "arguments[0].click();", opt);
                                    Console.WriteLine("✅ Đã click option PDF (tìm bằng text)");
                                    pdfClicked = true;
                                    break;
                                }
                            }
                        }
                        catch (Exception ex2)
                        {
                            Console.WriteLine($"⚠️ Không tìm thấy option PDF: {ex2.Message}");
                        }
                    }

                    // ===== 10. CHỜ FILE PDF =====
                    Console.WriteLine("Đang chờ file tải về...");
                    string latestPdfFile = null;
                    var deadline = DateTime.Now.AddSeconds(60);

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
                        await Task.Delay(500);
                    }

                    // ═══ NẾU LẦN 1 FAIL → THỬ CLICK LẠI ═══
                    if (latestPdfFile == null)
                    {
                        Console.WriteLine("⚠️ Lần 1 không có file — thử click lại PDF...");

                        try
                        {
                            var pdfRetry = driver.FindElement(By.CssSelector("div.dm-item.pdf.txt-download-pdf"));
                            ((IJavaScriptExecutor)driver).ExecuteScript(
                                "arguments[0].click();", pdfRetry);
                            Console.WriteLine("✅ Đã click lại PDF");
                        }
                        catch
                        {
                            try
                            {
                                ((IJavaScriptExecutor)driver).ExecuteScript(
                                    "document.querySelector('div.dm-item.pdf.txt-download-pdf')?.click();");
                            }
                            catch { }
                        }

                        // Chờ thêm 30s
                        var deadline2 = DateTime.Now.AddSeconds(30);
                        while (DateTime.Now < deadline2)
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
                            await Task.Delay(500);
                        }
                    }

                    if (latestPdfFile == null)
                    {
                        Console.WriteLine("❌ Không tải được file PDF hoặc quá thời gian chờ!");
                        XtraMessageBox.Show("Không tải được hóa đơn MISA");
                        return;
                    }

                    Console.WriteLine($"📄 File tải về: {latestPdfFile}");

                    // ===== 11. ĐỔI TÊN =====
                    string safeMst = MakeSafeFileName(mst);
                    string safeShdon = MakeSafeFileName(shdon);
                    string newPath = Path.Combine(baseDir, $"{safeMst}_{safeShdon}.pdf");

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

                    // ===== 12. MOVE FILE =====
                    File.Move(latestPdfFile, newPath);
                    Console.WriteLine($"✅ Đã lưu: {newPath}");

                    // ===== 13. DỌN RÁC =====
                    try
                    {
                        foreach (var f in Directory.GetFiles(tempDownloadFolder))
                        {
                            try { File.Delete(f); } catch { }
                        }
                    }
                    catch { }

                    // ===== 14. MỞ FILE =====
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
                        Console.WriteLine($"📂 Đã mở file: {newPath}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Lỗi: {ex.Message}");
                    XtraMessageBox.Show($"Lỗi tải MISA: {ex.Message}");
                }
                finally
                {
                    await Task.Delay(1500);

                    try { driver?.Quit(); } catch { }
                    try { driver?.Dispose(); } catch { }
                    Console.WriteLine("Đã đóng trình duyệt.");

                    try
                    {
                        if (Directory.Exists(tempDownloadFolder))
                            Directory.Delete(tempDownloadFolder, true);
                    }
                    catch { }
                }
            });
        }

        // ═══════════════════════════════════════════════════════════════
        // 🔐 XỬ LÝ CAPTCHA MISA — Chờ user nhập thủ công
        // ═══════════════════════════════════════════════════════════════
        private async Task HandleMisaCaptcha(IWebDriver driver, WebDriverWait wait)
        {
            try
            {
                await Task.Delay(1000);

                var captchaContainers = driver.FindElements(By.Id("captchaContainer"));

                if (captchaContainers.Count == 0)
                {
                    Console.WriteLine("ℹ️ Không có captcha");
                    return;
                }

                bool isVisible = false;
                try
                {
                    isVisible = captchaContainers[0].Displayed;
                    string style = captchaContainers[0].GetAttribute("style") ?? "";
                    if (style.Contains("display:none") || style.Contains("display: none"))
                        isVisible = false;
                }
                catch { }

                if (!isVisible)
                {
                    Console.WriteLine("ℹ️ Captcha ẩn — bỏ qua");
                    return;
                }

                Console.WriteLine("🔐 Captcha MISA xuất hiện — vui lòng nhập trong cửa sổ Chrome...");

                // ═══ ĐỢI USER NHẬP CAPTCHA (tối đa 90s) ═══
                var deadline = DateTime.Now.AddSeconds(90);

                while (DateTime.Now < deadline)
                {
                    try
                    {
                        var captchaInput = driver.FindElement(By.Id("captcha_captchaTextbox"));
                        string value = captchaInput.GetAttribute("value") ?? "";

                        if (!string.IsNullOrEmpty(value) && value.Length >= 4)
                        {
                            Console.WriteLine($"✅ User đã nhập captcha: '{value}'");
                            await Task.Delay(300);
                            return;
                        }
                    }
                    catch { }

                    await Task.Delay(500);
                }

                Console.WriteLine("⚠️ Hết thời gian chờ user nhập captcha");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ HandleMisaCaptcha: {ex.Message}");
            }
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
        private async Task TraCuuNCC(bool toggle, string mst, string mabm, string shdon, string savedPath, int type)
        {
            string year = dateEdit1.DateTime.Year.ToString();
            string month = dateEdit1.DateTime.Month.ToString();

            await Task.Run(async () =>
            {
                // ═══════════════════════════════════════════════════════
                // 1. XÁC ĐỊNH THƯ MỤC ĐÍCH
                // ═══════════════════════════════════════════════════════
                string baseDir;
                if (type == 1)
                    baseDir = Path.Combine(savedPath, $"HD{year}", "HDVao", month);
                else if (type == 2)
                    baseDir = Path.Combine(savedPath, $"HD{year}", "HDRa", month);
                else
                    baseDir = Path.Combine(savedPath, $"HD{year}", "HDKhac", month);

                Directory.CreateDirectory(baseDir);

                string tempDownloadFolder = Path.Combine(Path.GetTempPath(), "ViettelTemp", Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempDownloadFolder);

                // ═══════════════════════════════════════════════════════
                // ⚙️ CHROME OPTIONS
                // ✅ LƯU Ý: captcha slider Viettel rất khó chạy headless
                //    → nên bỏ --headless=new hoặc thêm --headless=old
                // ═══════════════════════════════════════════════════════
                ChromeOptions options = new ChromeOptions();
                // ⚠️ BỎ dòng dưới nếu captcha fail liên tục:
                // options.AddArgument("--headless=new");
                options.AddArgument("--window-size=1920,1080");
                options.AddArgument("--disable-features=OptimizationHints,MediaRouter,PrivacySandboxSettings4," +
                                    "PrivacySandboxAdsAPIs,InterestFeedContentSuggestions,Translate," +
                                    "AcceptCHFrame,ChromeLabs,CalculateNativeWinOcclusion," +
                                    "DialMediaRouteProvider,GlobalMediaControls," +
                                    "WebAuthentication,WebAuthenticationCable," +
                                    "WebBluetooth,WebBluetoothNewPermissionsBackend," +
                                    "WebUsb,WebSerial,WebHID,WebNFC," +
                                    "Serial,Bluetooth,HID,USB,NFC," +
                                    "FileSystemAccessAPI,IdleDetection," +
                                    "SyncDisabled,InfiniteSessionRestore");
                options.AddArgument("--disable-notifications");
                options.AddArgument("--disable-infobars");
                options.AddArgument("--disable-popup-blocking");
                options.AddArgument("--disable-save-password-bubble");
                options.AddArgument("--disable-translate");
                options.AddArgument("--no-first-run");
                options.AddArgument("--no-default-browser-check");
                options.AddArgument("--disable-default-apps");
                options.AddArgument("--disable-component-update");
                options.AddArgument("--disable-background-networking");
                options.AddArgument("--disable-sync");
                options.AddArgument("--metrics-recording-only");
                options.AddArgument("--disable-backgrounding-occluded-windows");
                options.AddArgument("--use-fake-ui-for-media-stream");
                options.AddArgument("--use-fake-device-for-media-stream");
                options.AddArgument("--deny-permission-prompts");

                options.AddArgument("--disable-gpu");
                options.AddArgument("--no-sandbox");
                options.AddArgument("--disable-dev-shm-usage");
                options.AddArgument("--disable-extensions");
                options.AddArgument("--disable-web-security");
                options.AddArgument("--disable-blink-features=AutomationControlled");
                options.AddExcludedArgument("enable-automation");
                options.AddAdditionalOption("useAutomationExtension", false);

                options.AddUserProfilePreference("download.default_directory", tempDownloadFolder);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("safebrowsing.enabled", true);
                options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);

                options.AddUserProfilePreference("profile.default_content_setting_values.media_stream", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.media_stream_mic", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.media_stream_camera", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.geolocation", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.notifications", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.midi_sysex", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.push_messaging", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.clipboard", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.sensors", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.payment_handler", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.usb_guard", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.bluetooth_guard", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.hid_guard", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.serial_guard", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.idle_detection", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.file_system_write_guard", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.protocol_handlers", 2);
                options.AddUserProfilePreference("profile.default_content_setting_values.native_file_system_write_guard", 2);
                options.AddUserProfilePreference("profile.managed_default_content_settings.images", 1);

                options.AddUserProfilePreference("credentials_enable_service", false);
                options.AddUserProfilePreference("profile.password_manager_enabled", false);
                options.AddUserProfilePreference("profile.password_manager_leak_detection", false);
                options.AddUserProfilePreference("autofill.profile_enabled", false);
                options.AddUserProfilePreference("autofill.credit_card_enabled", false);

                IWebDriver driver = null;
                try
                {
                    string driverDirectory = AppDomain.CurrentDomain.BaseDirectory;
                    var service = ChromeDriverService.CreateDefaultService(driverDirectory);
                    service.HideCommandPromptWindow = true;

                    driver = new ChromeDriver(service, options);
                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);

                    // ═══════════════════════════════════════════════════════
                    // 🛡️ CDP PROTOCOL
                    // ═══════════════════════════════════════════════════════
                    try
                    {
                        if (driver is ChromeDriver chromeDriver)
                        {
                            string[] perms = {
                        "notifications", "geolocation", "media", "midi",
                        "clipboard-read", "clipboard-write", "payment-handler",
                        "idle-detection", "periodic-background-sync"
                    };
                            foreach (var p in perms)
                            {
                                try
                                {
                                    chromeDriver.ExecuteCdpCommand(
                                        "Browser.setPermission",
                                        new Dictionary<string, object>
                                        {
                                    { "permission", new Dictionary<string, object> { { "name", p } } },
                                    { "setting", "denied" }
                                        });
                                }
                                catch { }
                            }

                            chromeDriver.ExecuteCdpCommand("Page.enable", new Dictionary<string, object>());
                            chromeDriver.ExecuteCdpCommand(
                                "Page.setBypassCSP",
                                new Dictionary<string, object> { { "enabled", true } });
                        }
                    }
                    catch (Exception cdpEx)
                    {
                        Console.WriteLine($"⚠️ CDP setup warning: {cdpEx.Message}");
                    }

                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                    // ═══════════════════════════════════════════════════════
                    // 2. TRUY CẬP TRANG
                    // ═══════════════════════════════════════════════════════
                    driver.Navigate().GoToUrl("https://vinvoice.viettel.vn/utilities/invoice-search");
                    await Task.Delay(500);
                    await TryDismissPermissionPopups(driver);

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
                    // 4. XỬ LÝ CAPTCHA — TỐI ĐA 5 LẦN
                    // ═══════════════════════════════════════════════════════
                    bool captchaSuccess = false;
                    int maxCaptchaRetries = 5;
                    int captchaAttempt = 0;

                    while (captchaAttempt < maxCaptchaRetries && !captchaSuccess)
                    {
                        captchaAttempt++;
                        Console.WriteLine($"🔵 Captcha lần {captchaAttempt}/{maxCaptchaRetries}");

                        try
                        {
                            await TryDismissPermissionPopups(driver);
                            await Task.Delay(500);

                            // ═══ DETECT CAPTCHA — kiểm tra 2 lần để tránh false negative ═══
                            bool DetectSlider() => driver.FindElements(By.CssSelector("div.slider-button")).Count > 0;
                            bool DetectRecaptcha() => driver.FindElements(By.CssSelector("iframe[src*='recaptcha'], div.g-recaptcha")).Count > 0;
                            bool DetectHcaptcha() => driver.FindElements(By.CssSelector("iframe[src*='hcaptcha'], div.h-captcha")).Count > 0;
                            bool DetectTurnstile() => driver.FindElements(By.CssSelector("iframe[src*='cloudflare'], div.cf-turnstile")).Count > 0;

                            bool hasSlider = DetectSlider();
                            bool hasRecaptcha = DetectRecaptcha();
                            bool hasHcaptcha = DetectHcaptcha();
                            bool hasTurnstile = DetectTurnstile();

                            // Kiểm tra lần 2 nếu chưa thấy gì
                            if (!hasSlider && !hasRecaptcha && !hasHcaptcha && !hasTurnstile)
                            {
                                await Task.Delay(800);
                                hasSlider = DetectSlider();
                                hasRecaptcha = DetectRecaptcha();
                                hasHcaptcha = DetectHcaptcha();
                                hasTurnstile = DetectTurnstile();
                            }

                            Console.WriteLine($"🔍 Captcha: slider={hasSlider}, recaptcha={hasRecaptcha}, hcaptcha={hasHcaptcha}, turnstile={hasTurnstile}");

                            if (!hasSlider && !hasRecaptcha && !hasHcaptcha && !hasTurnstile)
                            {
                                Console.WriteLine("✅ Không có captcha — tiếp tục");
                                captchaSuccess = true;
                                break;
                            }

                            if (hasRecaptcha || hasHcaptcha || hasTurnstile)
                            {
                                Console.WriteLine("⚠️ Captcha đổi loại — không xử lý được");
                                await Task.Delay(1000);
                                // ❌ KHÔNG set captchaSuccess = true → coi như fail
                                continue;
                            }

                            // ═══ XỬ LÝ SLIDER CAPTCHA ═══
                            if (hasSlider)
                            {
                                // Reset captcha trước khi xử lý (lần 2+)
                                if (captchaAttempt > 1)
                                {
                                    Console.WriteLine("🔄 Reset captcha + nhập lại form...");
                                    await ResetCaptchaAndForm(driver, mst, mabm);

                                    try
                                    {
                                        var waitCaptchaReady = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                                        waitCaptchaReady.Until(d =>
                                        {
                                            var pcs = d.FindElements(By.CssSelector("img.puzzle-piece"));
                                            var btns = d.FindElements(By.CssSelector("div.slider-button"));
                                            return (pcs.Count > 0 && btns.Count > 0) ? new object() : null;
                                        });
                                    }
                                    catch
                                    {
                                        Console.WriteLine("⚠️ Captcha chưa sẵn sàng — thử lại vòng sau");
                                        await Task.Delay(1000);
                                        continue;
                                    }

                                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                                }
                                else
                                {
                                    await EnsurePuzzleAtStart(driver);
                                }

                                // ═══ LẤY ẢNH CAPTCHA ═══
                                var bgImages = driver.FindElements(By.CssSelector("img.bg-image"));
                                if (bgImages.Count == 0)
                                {
                                    Console.WriteLine("⚠️ Không tìm thấy img.bg-image");
                                    await Task.Delay(500);
                                    continue;
                                }

                                var bgImage = bgImages[0];
                                string bgSrc = bgImage.GetAttribute("src");
                                if (string.IsNullOrEmpty(bgSrc) || bgSrc.Length < 100)
                                {
                                    Console.WriteLine("⚠️ bg-image src không hợp lệ");
                                    await Task.Delay(800);
                                    continue;
                                }

                                string bgPath = SaveBase64ToFile(bgSrc, "bg");

                                // ═══ KÍCH THƯỚC ẢNH ═══
                                int naturalW, naturalH;
                                using (var tmpBmp = new Bitmap(bgPath))
                                {
                                    naturalW = tmpBmp.Width;
                                    naturalH = tmpBmp.Height;
                                }
                                int displayedW = bgImage.Size.Width;
                                int displayedH = bgImage.Size.Height;

                                // ✅ FIX: dùng devicePixelRatio
                                double dpr = 1.0;
                                try
                                {
                                    var js = (IJavaScriptExecutor)driver;
                                    var dprObj = js.ExecuteScript("return window.devicePixelRatio || 1;");
                                    dpr = Convert.ToDouble(dprObj);
                                }
                                catch { }

                                double scaleX = (double)displayedW / naturalW;

                                Console.WriteLine($"📐 Gốc {naturalW}x{naturalH} | Hiển thị {displayedW}x{displayedH} | DPR={dpr:F2} | scaleX={scaleX:F3}");

                                int holeX_natural = FindHolePosition2(bgPath);
                                Console.WriteLine($"🎯 Ô đen (gốc): x={holeX_natural}px");

                                if (holeX_natural <= 0)
                                {
                                    Console.WriteLine("⚠️ Không tìm được vị trí ô đen");
                                    try { File.Delete(bgPath); } catch { }
                                    await Task.Delay(500);
                                    continue;
                                }

                                int holeX_displayed = (int)Math.Round(holeX_natural * scaleX);
                                Console.WriteLine($"🎯 Ô đen (hiển thị): x={holeX_displayed}px");

                                try { File.Delete(bgPath); } catch { }

                                // ═══ LẤY PUZZLE + SLIDER ═══
                                IWebElement puzzlePiece = null;
                                IWebElement sliderButton = null;

                                try
                                {
                                    var waitEl = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                                    waitEl.Until(d =>
                                    {
                                        var pcs = d.FindElements(By.CssSelector("img.puzzle-piece"));
                                        var btns = d.FindElements(By.CssSelector("div.slider-button"));

                                        if (pcs.Count > 0 && btns.Count > 0)
                                        {
                                            puzzlePiece = pcs[pcs.Count - 1];
                                            sliderButton = btns[btns.Count - 1];
                                            return new object();
                                        }
                                        return null;
                                    });
                                }
                                catch
                                {
                                    Console.WriteLine("⚠️ Timeout đợi puzzle/slider");
                                    await Task.Delay(800);
                                    continue;
                                }

                                if (puzzlePiece == null || sliderButton == null)
                                {
                                    Console.WriteLine("⚠️ Không lấy được puzzle/slider");
                                    await Task.Delay(800);
                                    continue;
                                }

                                ((IJavaScriptExecutor)driver).ExecuteScript(
                                    "arguments[0].scrollIntoView({block: 'center'});", sliderButton);
                                await Task.Delay(300);

                                // Kiểm tra puzzle ở vị trí đầu
                                string puzzleLeftStr = puzzlePiece.GetAttribute("style") ?? "";
                                Console.WriteLine($"🧩 Puzzle style: {puzzleLeftStr}");

                                if (!puzzleLeftStr.Contains("left: 0px") && !puzzleLeftStr.Contains("left:0px"))
                                {
                                    Console.WriteLine("⚠️ Puzzle không ở vị trí đầu — reset về 0");
                                    await EnsurePuzzleAtStart(driver);
                                    await Task.Delay(300);

                                    // Lấy lại element mới
                                    var pcs2 = driver.FindElements(By.CssSelector("img.puzzle-piece"));
                                    var btns2 = driver.FindElements(By.CssSelector("div.slider-button"));
                                    if (pcs2.Count == 0 || btns2.Count == 0) continue;
                                    puzzlePiece = pcs2[pcs2.Count - 1];
                                    sliderButton = btns2[btns2.Count - 1];
                                }

                                var puzzleSize = puzzlePiece.Size;
                                int puzzleW = puzzleSize.Width;
                                Console.WriteLine($"🧩 Puzzle-piece: {puzzleW}x{puzzleSize.Height}px");

                                // ═══ TÍNH KHOẢNG CÁCH KÉO ═══
                                int targetDistance = holeX_displayed - (puzzleW / 2);
                                Console.WriteLine($"📏 targetDistance={targetDistance}px");

                                if (targetDistance <= 5)
                                {
                                    Console.WriteLine("⚠️ Khoảng cách quá nhỏ");
                                    await Task.Delay(800);
                                    continue;
                                }

                                // ═══ KÉO SLIDER — DÙNG CDP ═══
                                await PerformSlideCdp(driver, sliderButton, targetDistance);
                                await Task.Delay(1500);

                                captchaSuccess = CheckCaptchaSuccess(driver);

                                // Debug log nếu fail
                                if (!captchaSuccess)
                                {
                                    try
                                    {
                                        var captchaInput = driver.FindElements(By.CssSelector("input[formcontrolname='recaptcha']"));
                                        if (captchaInput.Count > 0)
                                        {
                                            string val = captchaInput[0].GetAttribute("value") ?? "";
                                            string cls = captchaInput[0].GetAttribute("class") ?? "";
                                            Console.WriteLine($"🔍 DEBUG recaptcha: value='{val}', class='{cls}'");
                                        }

                                        var pieces = driver.FindElements(By.CssSelector("img.puzzle-piece"));
                                        if (pieces.Count > 0)
                                        {
                                            Console.WriteLine($"🔍 DEBUG puzzle style: {pieces[pieces.Count - 1].GetAttribute("style")}");
                                        }
                                    }
                                    catch { }
                                }

                                // ═══ RETRY TRONG CÙNG CAPTCHA ═══
                                if (!captchaSuccess)
                                {
                                    Console.WriteLine("⚠️ Fail lần 1 — thử kéo lại trong cùng captcha...");
                                    await Task.Delay(800);

                                    await EnsurePuzzleAtStart(driver);
                                    await Task.Delay(500);

                                    var pieces2 = driver.FindElements(By.CssSelector("img.puzzle-piece"));
                                    var btns2 = driver.FindElements(By.CssSelector("div.slider-button"));
                                    if (pieces2.Count > 0 && btns2.Count > 0)
                                    {
                                        var btn2 = btns2[btns2.Count - 1];
                                        var pz2 = pieces2[pieces2.Count - 1];
                                        int pzW2 = pz2.Size.Width;
                                        int target2 = holeX_displayed - (pzW2 / 2);

                                        await PerformSlideCdp(driver, btn2, target2);
                                        await Task.Delay(1500);

                                        captchaSuccess = CheckCaptchaSuccess(driver);
                                    }
                                }

                                // Debug screenshot
                                try
                                {
                                    string debugDir = Path.Combine(Path.GetTempPath(), "VinvoiceDebug");
                                    Directory.CreateDirectory(debugDir);
                                    string shot = Path.Combine(debugDir, $"slide_{captchaAttempt}.png");
                                    ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile(shot);
                                    Console.WriteLine($"📸 Debug: {shot}");
                                }
                                catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Lỗi captcha lần {captchaAttempt}: {ex.Message}");
                        }

                        if (!captchaSuccess)
                        {
                            await Task.Delay(500);
                        }
                    }

                    if (!captchaSuccess)
                    {
                        Console.WriteLine("❌ Thử captcha nhiều lần không thành công.");
                        return;
                    }

                    Console.WriteLine("✅ Captcha OK — chuẩn bị tìm kiếm");
                    await Task.Delay(500);

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
                        await Task.Delay(500);

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

                    for (int retry = 0; retry < 3; retry++)
                    {
                        try
                        {
                            if (Directory.Exists(tempDownloadFolder))
                            {
                                await Task.Delay(500);
                                Directory.Delete(tempDownloadFolder, true);
                                Console.WriteLine($"🗑️ Đã xóa folder tạm: {tempDownloadFolder}");
                            }
                            break;
                        }
                        catch (IOException ioEx)
                        {
                            Console.WriteLine($"⚠️ Xóa folder lần {retry + 1} bị lock: {ioEx.Message}");
                            await Task.Delay(1000);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Lỗi xóa folder: {ex.Message}");
                            break;
                        }
                    }
                }
            });
        }
        private async Task PerformSlideCdp(IWebDriver driver, IWebElement sliderButton, int targetDistance)
        {
            try
            {
                if (!(driver is ChromeDriver chromeDriver))
                {
                    Console.WriteLine("⚠️ Không phải ChromeDriver — fallback Actions");
                    await PerformSlideWithActions(driver, sliderButton, targetDistance);
                    return;
                }

                var js = (IJavaScriptExecutor)driver;

                // ═══ LẤY TỌA ĐỘ TÂM SLIDER TRONG VIEWPORT ═══
                var rect = (Dictionary<string, object>)js.ExecuteScript(@"
            var r = arguments[0].getBoundingClientRect();
            return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
        ", sliderButton);

                int startX = Convert.ToInt32(rect["x"]);
                int startY = Convert.ToInt32(rect["y"]);

                Console.WriteLine($"🎯 CDP slide: start=({startX},{startY}) distance={targetDistance}");

                // ═══ MOUSE DOWN ═══
                chromeDriver.ExecuteCdpCommand("Input.dispatchMouseEvent", new Dictionary<string, object>
        {
            { "type", "mousePressed" },
            { "x", startX },
            { "y", startY },
            { "button", "left" },
            { "buttons", 1 },
            { "clickCount", 1 }
        });

                await Task.Delay(30);

                // ═══ DI CHUYỂN NHIỀU BƯỚC ═══
                var rnd = new Random();
                int steps = 18 + rnd.Next(8);
                double stepDist = (double)targetDistance / steps;

                for (int i = 1; i <= steps; i++)
                {
                    int cx = startX + (int)Math.Round(stepDist * i);
                    int cy = startY + rnd.Next(-1, 2);

                    chromeDriver.ExecuteCdpCommand("Input.dispatchMouseEvent", new Dictionary<string, object>
            {
                { "type", "mouseMoved" },
                { "x", cx },
                { "y", cy },
                { "button", "left" },
                { "buttons", 1 }
            });

                    await Task.Delay(rnd.Next(8, 20));
                }

                // ═══ ĐẢM BẢO ĐẾN ĐÚNG VỊ TRÍ ═══
                int finalX = startX + targetDistance;
                chromeDriver.ExecuteCdpCommand("Input.dispatchMouseEvent", new Dictionary<string, object>
        {
            { "type", "mouseMoved" },
            { "x", finalX },
            { "y", startY },
            { "button", "left" },
            { "buttons", 1 }
        });

                await Task.Delay(50);

                // ═══ MOUSE UP ═══
                chromeDriver.ExecuteCdpCommand("Input.dispatchMouseEvent", new Dictionary<string, object>
        {
            { "type", "mouseReleased" },
            { "x", finalX },
            { "y", startY },
            { "button", "left" },
            { "buttons", 0 },
            { "clickCount", 1 }
        });

                Console.WriteLine($"✅ CDP slide hoàn tất — kéo {targetDistance}px trong {steps} bước");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ CDP slide lỗi: {ex.Message} — fallback Actions");
                await PerformSlideWithActions(driver, sliderButton, targetDistance);
            }
        }
        // ═══════════════════════════════════════════════════════════════
        // 🎬 HELPER: THỰC HIỆN KÉO SLIDER
        // ═══════════════════════════════════════════════════════════════
        private async Task PerformSlide(IWebDriver driver, IWebElement sliderButton, int targetDistance)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;

                // Tìm container cha để dispatch event đúng
                var puzzlePieces = driver.FindElements(By.CssSelector("img.puzzle-piece"));
                var sliderBars = driver.FindElements(By.CssSelector("div.slider-bar"));

                if (puzzlePieces.Count == 0 || sliderBars.Count == 0)
                {
                    Console.WriteLine("⚠️ Thiếu puzzle/slider-bar — dùng fallback Actions");
                    await PerformSlideWithActions(driver, sliderButton, targetDistance);
                    return;
                }

                var puzzlePiece = puzzlePieces[puzzlePieces.Count - 1];
                var sliderBar = sliderBars[sliderBars.Count - 1];

                // ═══ DÙNG JS DISPATCH — NHANH HƠN 5-10 LẦN ═══
                // Mô phỏng chuỗi event: mousedown → mousemove (nhiều bước) → mouseup
                string jsScript = @"
            (function(btn, bar, puzzle, distance) {
                try {
                    // Lấy tọa độ trung tâm button
                    var btnRect = btn.getBoundingClientRect();
                    var barRect = bar.getBoundingClientRect();
                    var startX = btnRect.left + btnRect.width / 2;
                    var startY = btnRect.top + btnRect.height / 2;

                    // Số bước di chuyển — càng nhiều càng mượt
                    var steps = 25;
                    var stepDist = distance / steps;

                    // Helper tạo event
                    function fireEvent(el, type, x, y, buttons) {
                        var ev = new MouseEvent(type, {
                            bubbles: true,
                            cancelable: true,
                            view: window,
                            clientX: x,
                            clientY: y,
                            button: 0,
                            buttons: buttons
                        });
                        el.dispatchEvent(ev);
                    }

                    // 1. mousedown trên slider button
                    fireEvent(btn, 'mousedown', startX, startY, 1);

                    // 2. mousemove liên tiếp
                    var currentX = startX;
                    var currentY = startY;

                    for (var i = 0; i < steps; i++) {
                        currentX += stepDist;
                        // Jitter nhỏ
                        currentY = startY + (Math.random() * 2 - 1);

                        // Dispatch trên document (vì Angular thường listen ở document/window)
                        fireEvent(document, 'mousemove', currentX, currentY, 1);
                        fireEvent(bar, 'mousemove', currentX, currentY, 1);
                    }

                    // 3. mouseup
                    fireEvent(document, 'mouseup', currentX, currentY, 0);
                    fireEvent(btn, 'mouseup', currentX, currentY, 0);

                    return true;
                } catch(e) {
                    return 'ERROR: ' + e.message;
                }
            })(arguments[0], arguments[1], arguments[2], arguments[3]);
        ";

                var result = js.ExecuteScript(jsScript, sliderButton, sliderBar, puzzlePiece, targetDistance);
                Console.WriteLine($"⚡ JS dispatch result: {result}");

                // Vì JS chạy tức thì, đợi 1 nhịp cho Angular cập nhật DOM
                await Task.Delay(300);

                // Kiểm tra puzzle đã di chuyển chưa
                try
                {
                    var pieces = driver.FindElements(By.CssSelector("img.puzzle-piece"));
                    if (pieces.Count > 0)
                    {
                        string style = pieces[pieces.Count - 1].GetAttribute("style") ?? "";
                        Console.WriteLine($"🧩 Puzzle sau kéo: {style}");

                        // Nếu puzzle vẫn ở left: 0 → JS không có tác dụng → fallback Actions
                        if (style.Contains("left: 0px") || style.Contains("left:0px"))
                        {
                            Console.WriteLine("⚠️ Puzzle không di chuyển — dùng fallback Actions");
                            await PerformSlideWithActions(driver, sliderButton, targetDistance);
                        }
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ JS slide lỗi: {ex.Message} — fallback Actions");
                await PerformSlideWithActions(driver, sliderButton, targetDistance);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // 🎬 FALLBACK: Kéo bằng Actions (chậm hơn nhưng chắc chắn)
        // ═══════════════════════════════════════════════════════════════
        private async Task PerformSlideWithActions(IWebDriver driver, IWebElement sliderButton, int targetDistance)
        {
            Actions actions = new Actions(driver);
            var rnd = new Random();

            actions.MoveToElement(sliderButton).Perform();
            await Task.Delay(30);

            actions.ClickAndHold(sliderButton).Perform();
            await Task.Delay(30);

            // Giảm số bước xuống 8-12 để nhanh hơn
            int steps = 8 + rnd.Next(5);
            double stepDist = (double)targetDistance / steps;

            for (int i = 0; i < steps; i++)
            {
                actions.MoveByOffset((int)Math.Round(stepDist), rnd.Next(-1, 2)).Perform();
                await Task.Delay(rnd.Next(3, 7));
            }

            double moved = stepDist * steps;
            double residual = targetDistance - moved;
            if (Math.Abs(residual) > 0.5)
                actions.MoveByOffset((int)Math.Round(residual), 0).Perform();

            await Task.Delay(30);
            actions.Release().Perform();

            Console.WriteLine($"✅ Fallback Actions: kéo {targetDistance}px với {steps} bước");
        }

        // ═══════════════════════════════════════════════════════════════
        // 🔄 HELPER: ĐẢM BẢO PUZZLE + SLIDER VỀ VỊ TRÍ 0
        // ═══════════════════════════════════════════════════════════════
        private async Task EnsurePuzzleAtStart(IWebDriver driver)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;

                var pieces = driver.FindElements(By.CssSelector("img.puzzle-piece"));
                foreach (var p in pieces)
                {
                    try
                    {
                        js.ExecuteScript("arguments[0].style.left = '0px';", p);
                    }
                    catch { }
                }

                var btns = driver.FindElements(By.CssSelector("div.slider-button"));
                foreach (var b in btns)
                {
                    try
                    {
                        js.ExecuteScript("arguments[0].style.left = '0px';", b);
                    }
                    catch { }
                }

                var progresses = driver.FindElements(By.CssSelector("div.slider-progress"));
                foreach (var pr in progresses)
                {
                    try
                    {
                        js.ExecuteScript("arguments[0].style.width = '0px';", pr);
                    }
                    catch { }
                }

                Console.WriteLine("✅ Đã reset puzzle + slider về vị trí 0");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ EnsurePuzzleAtStart: {ex.Message}");
            }
            await Task.Delay(300);
        }

        // ═══════════════════════════════════════════════════════════════
        // 🔄 HELPER: RESET CAPTCHA + NHẬP LẠI FORM
        // ═══════════════════════════════════════════════════════════════
        private async Task ResetCaptchaAndForm(IWebDriver driver, string mst, string mabm)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;

                // ═══ BƯỚC 1: Thử click nút reload captcha (không cần refresh trang) ═══
                var reloadBtns = driver.FindElements(By.CssSelector("span.reload"));
                bool reloaded = false;

                foreach (var rb in reloadBtns)
                {
                    try
                    {
                        if (rb.Displayed && rb.Enabled)
                        {
                            js.ExecuteScript("arguments[0].click();", rb);
                            Console.WriteLine("✅ Đã click nút reload captcha");
                            reloaded = true;
                            break;
                        }
                    }
                    catch { }
                }

                if (reloaded)
                {
                    // Đợi ảnh captcha mới load xong (src thay đổi)
                    await Task.Delay(1500);

                    // Đợi puzzle-piece xuất hiện lại
                    try
                    {
                        var waitPuzzle = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        waitPuzzle.Until(d =>
                        {
                            var pcs = d.FindElements(By.CssSelector("img.puzzle-piece"));
                            return pcs.Count > 0 ? pcs[0] : null;
                        });
                        Console.WriteLine("✅ Puzzle-piece đã render lại sau reload");
                    }
                    catch
                    {
                        Console.WriteLine("⚠️ Không thấy puzzle-piece sau reload");
                    }

                    await EnsurePuzzleAtStart(driver);
                    await Task.Delay(300);
                    return;
                }

                // ═══ BƯỚC 2: Không có nút reload → refresh cả trang ═══
                Console.WriteLine("🔄 Không có nút reload — refresh trang...");
                driver.Navigate().Refresh();

                // Đợi trang load lại (ít nhất 1 element form xuất hiện)
                await Task.Delay(2500);
                await TryDismissPermissionPopups(driver);

                // ═══ BƯỚC 3: Đợi form MST xuất hiện rồi nhập lại ═══
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                try
                {
                    var taxRetry = wait.Until(d =>
                    {
                        try
                        {
                            var el = d.FindElement(By.CssSelector("input[formcontrolname='supplierTaxCode']"));
                            return (el != null && el.Displayed) ? el : null;
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
                            return (el != null && el.Displayed) ? el : null;
                        }
                        catch { return null; }
                    });
                    codeRetry.Clear();
                    codeRetry.SendKeys(mabm);

                    Console.WriteLine("✅ Đã nhập lại MST + mã sau refresh");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Không nhập lại được form: {ex.Message}");
                }

                // ═══ BƯỚC 4: ĐỢI CAPTCHA RENDER XONG (đây là fix chính) ═══
                Console.WriteLine("⏳ Đợi captcha render lại...");

                try
                {
                    var waitCaptcha = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                    // Đợi đồng thời cả puzzle-piece VÀ slider-button xuất hiện
                    waitCaptcha.Until(d =>
                    {
                        var pcs = d.FindElements(By.CssSelector("img.puzzle-piece"));
                        var btns = d.FindElements(By.CssSelector("div.slider-button"));
                        var bgs = d.FindElements(By.CssSelector("img.bg-image"));

                        bool ok = pcs.Count > 0 && btns.Count > 0 && bgs.Count > 0;

                        if (ok)
                        {
                            // Kiểm tra ảnh bg-image có src hợp lệ (không phải rỗng)
                            try
                            {
                                string src = bgs[bgs.Count - 1].GetAttribute("src");
                                if (string.IsNullOrEmpty(src) || src.Length < 100)
                                    return null;
                            }
                            catch { return null; }
                        }

                        return ok ? new object() : null;
                    });

                    Console.WriteLine("✅ Captcha đã render xong");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Timeout đợi captcha: {ex.Message}");
                }

                // Đợi thêm 1 nhịp cho Angular ổn định DOM
                await Task.Delay(500);
                await EnsurePuzzleAtStart(driver);
                await Task.Delay(300);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ ResetCaptchaAndForm: {ex.Message}");
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // 🔧 HELPER: ĐÓNG POPUP PERMISSION + OVERRIDE API
        // ═══════════════════════════════════════════════════════════════
        private async Task TryDismissPermissionPopups(IWebDriver driver)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;

                js.ExecuteScript(@"
            (function() {
                try {
                    window.alert = function() { return true; };
                    window.confirm = function() { return true; };
                    window.prompt = function() { return null; };
                    window.print = function() { };
                    window.open = function() { return null; };
                } catch(e) {}
            })();
        ");

                js.ExecuteScript(@"
            (function() {
                try {
                    if (navigator.permissions && navigator.permissions.query) {
                        navigator.permissions.query = function(desc) {
                            return Promise.resolve({
                                state: 'denied', onchange: null,
                                addEventListener: function() {},
                                removeEventListener: function() {}
                            });
                        };
                    }
                    if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
                        navigator.mediaDevices.getUserMedia = function() {
                            return Promise.reject(new Error('Permission denied'));
                        };
                    }
                    if (navigator.mediaDevices && navigator.mediaDevices.getDisplayMedia) {
                        navigator.mediaDevices.getDisplayMedia = function() {
                            return Promise.reject(new Error('Permission denied'));
                        };
                    }
                    if (navigator.geolocation) {
                        navigator.geolocation.getCurrentPosition = function(s, e) {
                            if (e) e({ code: 1, message: 'User denied Geolocation' });
                        };
                        navigator.geolocation.watchPosition = function() { return 0; };
                    }
                    if (navigator.bluetooth && navigator.bluetooth.requestDevice) {
                        navigator.bluetooth.requestDevice = function() {
                            return Promise.reject(new Error('Permission denied'));
                        };
                    }
                    if (navigator.usb && navigator.usb.requestDevice) {
                        navigator.usb.requestDevice = function() {
                            return Promise.reject(new Error('Permission denied'));
                        };
                    }
                    if (navigator.serial && navigator.serial.requestPort) {
                        navigator.serial.requestPort = function() {
                            return Promise.reject(new Error('Permission denied'));
                        };
                    }
                    if (navigator.hid && navigator.hid.requestDevice) {
                        navigator.hid.requestDevice = function() {
                            return Promise.reject(new Error('Permission denied'));
                        };
                    }
                    if (window.showOpenFilePicker) {
                        window.showOpenFilePicker = function() {
                            return Promise.reject(new Error('Permission denied'));
                        };
                    }
                    if (window.showSaveFilePicker) {
                        window.showSaveFilePicker = function() {
                            return Promise.reject(new Error('Permission denied'));
                        };
                    }
                    if (window.showDirectoryPicker) {
                        window.showDirectoryPicker = function() {
                            return Promise.reject(new Error('Permission denied'));
                        };
                    }
                } catch(e) {}
            })();
        ");

                try
                {
                    var blockBtns = driver.FindElements(By.XPath(
                        "//button[contains(translate(., 'BLOCK', 'block'), 'block') " +
                        "or contains(., 'Chặn') or contains(., 'Từ chối') " +
                        "or contains(., 'Deny') or contains(., 'Cancel')]"));
                    foreach (var btn in blockBtns)
                    {
                        try
                        {
                            if (btn.Displayed && btn.Enabled)
                            {
                                btn.Click();
                                await Task.Delay(200);
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                await Task.Delay(200);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ TryDismissPermissionPopups: {ex.Message}");
            }
        }

        /// <summary>
        /// Tìm tâm ô đen trong ảnh captcha Viettel.
        /// Pipeline: grayscale → adaptive threshold → morphology → connected components → scoring.
        /// </summary>
        static int FindHolePosition2(string imagePath)
        {
            try
            {
                if (!File.Exists(imagePath)) return -1;

                using (Bitmap bmp = new Bitmap(imagePath))
                {
                    int width = bmp.Width;
                    int height = bmp.Height;

                    // ═══════════════════════════════════════════════════════
                    // BƯỚC 1: Lấy grayscale + histogram
                    // ═══════════════════════════════════════════════════════
                    var rect = new System.Drawing.Rectangle(0, 0, width, height);
                    BitmapData bmpData = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                    int bytesPerPixel = 3;
                    int stride = bmpData.Stride;
                    IntPtr ptr = bmpData.Scan0;
                    int bytes = Math.Abs(stride) * height;
                    byte[] rgb = new byte[bytes];
                    Marshal.Copy(ptr, rgb, 0, bytes);
                    bmp.UnlockBits(bmpData);

                    byte[] grayMap = new byte[width * height];
                    int[] hist = new int[256];

                    for (int y = 0; y < height; y++)
                    {
                        int rowOff = y * stride;
                        for (int x = 0; x < width; x++)
                        {
                            int idx = rowOff + x * bytesPerPixel;
                            byte b = rgb[idx], g = rgb[idx + 1], r = rgb[idx + 2];
                            byte gray = (byte)((r * 299 + g * 587 + b * 114) / 1000);
                            grayMap[y * width + x] = gray;
                            hist[gray]++;
                        }
                    }

                    // ═══════════════════════════════════════════════════════
                    // BƯỚC 2: Adaptive threshold — 3% darkest pixels
                    // ═══════════════════════════════════════════════════════
                    int totalPixels = width * height;
                    int targetDarkPixels = (int)(totalPixels * 0.03);

                    int cumulative = 0;
                    int adaptiveThr = 25;
                    for (int i = 0; i < 256; i++)
                    {
                        cumulative += hist[i];
                        if (cumulative >= targetDarkPixels)
                        {
                            adaptiveThr = i;
                            break;
                        }
                    }
                    adaptiveThr = Math.Max(15, Math.Min(adaptiveThr, 60));
                    Console.WriteLine($"📊 adaptiveThr={adaptiveThr} (3% darkest)");

                    // ═══════════════════════════════════════════════════════
                    // BƯỚC 3: Morphology erode + dilate
                    // ═══════════════════════════════════════════════════════
                    byte[] binary = new byte[width * height];
                    for (int i = 0; i < grayMap.Length; i++)
                        binary[i] = grayMap[i] < adaptiveThr ? (byte)1 : (byte)0;

                    byte[] eroded = MorphErode(binary, width, height);
                    byte[] morphed = MorphDilate(eroded, width, height);

                    // ═══════════════════════════════════════════════════════
                    // BƯỚC 4: Connected components
                    // ═══════════════════════════════════════════════════════
                    var comps = FindComponentsAdaptive(morphed, width, height);
                    Console.WriteLine($"🔢 {comps.Count} components");

                    if (comps.Count == 0)
                    {
                        Console.WriteLine("⚠️ Không có component — fallback");
                        return FallbackColumnScan(grayMap, width, height, adaptiveThr);
                    }

                    // ═══════════════════════════════════════════════════════
                    // BƯỚC 5: Scoring
                    // ═══════════════════════════════════════════════════════
                    int expectedSize = (int)(width * 0.13);
                    Console.WriteLine($"📐 Ảnh {width}x{height}, expectedSize~{expectedSize}px");

                    var valid = new List<(int cx, int cy, double score, string reason)>();

                    foreach (var c in comps)
                    {
                        // Lọc cứng
                        double ratio = (double)c.w / c.h;
                        if (ratio < 0.80 || ratio > 1.25) continue;

                        if (c.w < expectedSize * 0.6 || c.w > expectedSize * 1.5) continue;
                        if (c.h < expectedSize * 0.6 || c.h > expectedSize * 1.5) continue;

                        if (c.density < 0.90) continue;

                        if (c.cx < width * 0.05 || c.cx > width * 0.95) continue;
                        if (c.cy < height * 0.05 || c.cy > height * 0.95) continue;

                        double dyFromCenter = Math.Abs(c.cy - height / 2.0) / height;
                        if (dyFromCenter > 0.35) continue;

                        // Verify độ tối thực
                        double avgGray = GetComponentAvgGray(grayMap, width, c);
                        if (avgGray > 40) continue;

                        // Score
                        double sizeScore = 1.0 - Math.Abs(c.w - expectedSize) / (double)expectedSize;
                        double centerYScore = 1.0 - dyFromCenter;
                        double densityScore = c.density;
                        double darknessScore = 1.0 - avgGray / 60.0;
                        double posXScore = 1.0 - Math.Abs(c.cx - width * 0.5) / (width * 0.5);

                        double score =
                            sizeScore * 3.0 +
                            centerYScore * 2.0 +
                            densityScore * 2.5 +
                            darknessScore * 2.0 +
                            posXScore * 1.0;

                        valid.Add((c.cx, c.cy, score,
                            $"size={c.w}x{c.h} cy={c.cy} dens={c.density:F3} gray={avgGray:F1}"));
                    }

                    if (valid.Count == 0)
                    {
                        Console.WriteLine("⚠️ Không candidate nào qua lọc — fallback");
                        return FallbackColumnScan(grayMap, width, height, adaptiveThr);
                    }

                    valid.Sort((a, b) => b.score.CompareTo(a.score));

                    Console.WriteLine($"📊 Top {Math.Min(5, valid.Count)}:");
                    for (int i = 0; i < Math.Min(5, valid.Count); i++)
                        Console.WriteLine($"   #{i + 1}: x={valid[i].cx} score={valid[i].score:F2} {valid[i].reason}");

                    // Chọn
                    if (valid.Count >= 2 && valid[0].score > valid[1].score * 1.25)
                    {
                        Console.WriteLine($"🎯 Top 1 vượt trội: x={valid[0].cx}");
                        return valid[0].cx;
                    }

                    if (valid.Count >= 2 && Math.Abs(valid[0].cx - valid[1].cx) < 25)
                    {
                        int avgX = (valid[0].cx + valid[1].cx) / 2;
                        Console.WriteLine($"🎯 Top 1+2 gần: avg x={avgX}");
                        return avgX;
                    }

                    Console.WriteLine($"🎯 Chọn top 1: x={valid[0].cx}");
                    return valid[0].cx;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ FindHolePosition2: {ex.Message}");
                return -1;
            }
        }
        static int FallbackColumnScan(byte[] grayMap, int width, int height, int threshold)
        {
            int marginX = (int)(width * 0.15);
            int marginY = (int)(height * 0.15);

            int[] colDensity = new int[width];
            for (int y = marginY; y < height - marginY; y++)
            {
                int rowOff = y * width;
                for (int x = marginX; x < width - marginX; x++)
                {
                    if (grayMap[rowOff + x] <= threshold)
                        colDensity[x]++;
                }
            }

            int bestCol = -1, bestVal = 0;
            for (int x = marginX; x < width - marginX; x++)
            {
                if (colDensity[x] > bestVal)
                {
                    bestVal = colDensity[x];
                    bestCol = x;
                }
            }

            if (bestCol < 0 || bestVal < height * 0.05)
            {
                Console.WriteLine("⚠️ Fallback cũng thất bại");
                return -1;
            }

            Console.WriteLine($"🎯 Fallback column: x={bestCol}, density={bestVal}");
            return bestCol;
        }
        // ═══════════════════════════════════════════════════════════════
        // Component struct
        // ═══════════════════════════════════════════════════════════════
        public struct Component
        {
            public int minX, maxX, minY, maxY;
            public int cx, cy, w, h;
            public int area;
            public double density;
        }

        // ═══════════════════════════════════════════════════════════════
        // Morphology
        // ═══════════════════════════════════════════════════════════════
        static byte[] MorphErode(byte[] src, int w, int h)
        {
            byte[] dst = new byte[w * h];
            for (int y = 1; y < h - 1; y++)
                for (int x = 1; x < w - 1; x++)
                {
                    bool allOn = true;
                    for (int dy = -1; dy <= 1 && allOn; dy++)
                        for (int dx = -1; dx <= 1; dx++)
                            if (src[(y + dy) * w + (x + dx)] == 0) { allOn = false; break; }
                    dst[y * w + x] = allOn ? (byte)1 : (byte)0;
                }
            return dst;
        }

        static byte[] MorphDilate(byte[] src, int w, int h)
        {
            byte[] dst = new byte[w * h];
            for (int y = 1; y < h - 1; y++)
                for (int x = 1; x < w - 1; x++)
                {
                    bool anyOn = false;
                    for (int dy = -1; dy <= 1 && !anyOn; dy++)
                        for (int dx = -1; dx <= 1; dx++)
                            if (src[(y + dy) * w + (x + dx)] != 0) { anyOn = true; break; }
                    dst[y * w + x] = anyOn ? (byte)1 : (byte)0;
                }
            return dst;
        }

        // ═══════════════════════════════════════════════════════════════
        // Connected components (8-hướng, dùng Stack)
        // ═══════════════════════════════════════════════════════════════
        static List<Component> FindComponentsAdaptive(byte[] binaryMap, int width, int height)
        {
            var result = new List<Component>();
            int totalPixels = width * height;
            bool[] visited = new bool[totalPixels];

            int[] dx = { -1, 1, 0, 0, -1, -1, 1, 1 };
            int[] dy = { 0, 0, -1, 1, -1, 1, -1, 1 };

            var stack = new Stack<int>(1024);

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    int p = y * width + x;
                    if (visited[p] || binaryMap[p] == 0) continue;

                    int area = 0;
                    int minX = int.MaxValue, maxX = int.MinValue;
                    int minY = int.MaxValue, maxY = int.MinValue;

                    stack.Clear();
                    stack.Push(p);
                    visited[p] = true;

                    while (stack.Count > 0)
                    {
                        int cur = stack.Pop();
                        int cy = cur / width;
                        int cx = cur - cy * width;

                        area++;
                        if (cx < minX) minX = cx;
                        if (cx > maxX) maxX = cx;
                        if (cy < minY) minY = cy;
                        if (cy > maxY) maxY = cy;

                        for (int d = 0; d < 8; d++)
                        {
                            int nx = cx + dx[d];
                            int ny = cy + dy[d];
                            if ((uint)nx >= (uint)width || (uint)ny >= (uint)height) continue;
                            int np = ny * width + nx;
                            if (visited[np] || binaryMap[np] == 0) continue;
                            visited[np] = true;
                            stack.Push(np);
                        }
                    }

                    if (area < 30) continue;
                    if (area > totalPixels * 0.4) continue;

                    int compW = maxX - minX + 1;
                    int compH = maxY - minY + 1;
                    int boundingArea = compW * compH;
                    if (boundingArea == 0) continue;

                    double density = (double)area / boundingArea;
                    if (density < 0.70) continue;

                    double aspect = (double)compW / compH;
                    if (aspect < 0.55 || aspect > 1.80) continue;

                    result.Add(new Component
                    {
                        minX = minX,
                        maxX = maxX,
                        minY = minY,
                        maxY = maxY,
                        cx = (minX + maxX) / 2,
                        cy = (minY + maxY) / 2,
                        w = compW,
                        h = compH,
                        area = area,
                        density = density
                    });
                }
            }

            result.Sort((a, b) => b.area.CompareTo(a.area));
            return result;
        }

        static double GetComponentAvgGray(byte[] gray, int w, Component c)
        {
            long sum = 0; int count = 0;
            for (int y = c.minY; y <= c.maxY; y++)
                for (int x = c.minX; x <= c.maxX; x++)
                {
                    sum += gray[y * w + x];
                    count++;
                }
            return count > 0 ? (double)sum / count : 255;
        }
        private bool CheckCaptchaSuccess(IWebDriver driver)
        {
            try
            {
                // ═══ 1. Kiểm tra thông báo lỗi ═══
                string[] errorSelectors = new[]
                {
            "div.error-msg",
            "div.ant-message-error",
            "div.ant-notification-notice-error",
            "div.toast-error",
            "div.alert-danger",
            "jhi-error-messages div.error-msg",
            ".ant-form-item-explain-error"
        };

                string[] errorKeywords = new[]
                {
            "captcha không đúng",
            "captcha sai",
            "sai captcha",
            "vui lòng xác thực lại",
            "kéo chính xác",
            "captcha không chính xác",
            "không đúng, vui lòng thử lại",
            "sai, vui lòng thử lại",
            "vui lòng thử lại"
        };

                foreach (var sel in errorSelectors)
                {
                    try
                    {
                        var errors = driver.FindElements(By.CssSelector(sel));
                        foreach (var err in errors)
                        {
                            if (!err.Displayed) continue;
                            string txt = (err.Text ?? "").ToLower();

                            // ✅ Chỉ bỏ qua khi message là validate form gốc (chưa submit)
                            // Kiểm tra input recaptcha có value không
                            bool hasRecaptchaValue = false;
                            try
                            {
                                var hidden = driver.FindElements(By.CssSelector("input[formcontrolname='recaptcha']"));
                                if (hidden.Count > 0)
                                {
                                    string v = hidden[0].GetAttribute("value") ?? "";
                                    hasRecaptchaValue = !string.IsNullOrEmpty(v);
                                }
                            }
                            catch { }

                            // Nếu chưa có value → đây là validate form, bỏ qua
                            if (!hasRecaptchaValue && txt.Contains("vui lòng xác thực captcha"))
                                continue;

                            foreach (var kw in errorKeywords)
                            {
                                if (txt.Contains(kw))
                                {
                                    Console.WriteLine($"❌ Phát hiện lỗi captcha: '{err.Text}'");
                                    return false;
                                }
                            }
                        }
                    }
                    catch { }
                }

                // ═══ 2. Toast toàn cục ═══
                try
                {
                    var messages = driver.FindElements(By.CssSelector(
                        ".ant-message-notice-content, .ant-notification-notice-message, .ant-notification-notice-description"));
                    foreach (var msg in messages)
                    {
                        if (!msg.Displayed) continue;
                        string txt = (msg.Text ?? "").ToLower();
                        foreach (var kw in errorKeywords)
                        {
                            if (txt.Contains(kw))
                            {
                                Console.WriteLine($"❌ Toast lỗi: '{msg.Text}'");
                                return false;
                            }
                        }
                    }
                }
                catch { }

                // ═══ 3. Puzzle/slider reset về 0 → SAI ═══
                try
                {
                    var pieces = driver.FindElements(By.CssSelector("img.puzzle-piece"));
                    var btns = driver.FindElements(By.CssSelector("div.slider-button"));

                    if (pieces.Count == 0 && btns.Count == 0)
                    {
                        Console.WriteLine("✅ Captcha biến mất — ĐÚNG");
                        return true;
                    }

                    if (pieces.Count > 0)
                    {
                        var style = pieces[pieces.Count - 1].GetAttribute("style") ?? "";
                        if (style.Contains("left: 0px") || style.Contains("left:0px"))
                        {
                            System.Threading.Thread.Sleep(500);
                            style = pieces[pieces.Count - 1].GetAttribute("style") ?? "";
                            if (style.Contains("left: 0px") || style.Contains("left:0px"))
                            {
                                Console.WriteLine("❌ Puzzle về 0 — SAI");
                                return false;
                            }
                        }
                    }
                }
                catch { }

                // ═══ 4. recaptcha input có value → ĐÚNG ═══
                try
                {
                    var captchaInput = driver.FindElements(By.CssSelector("input[formcontrolname='recaptcha']"));
                    if (captchaInput.Count > 0)
                    {
                        string val = captchaInput[0].GetAttribute("value") ?? "";
                        string cls = captchaInput[0].GetAttribute("class") ?? "";

                        if (!string.IsNullOrEmpty(val))
                        {
                            Console.WriteLine($"✅ recaptcha value='{val}' — ĐÚNG");
                            return true;
                        }

                        if (cls.Contains("ng-valid") && !cls.Contains("ng-invalid"))
                        {
                            Console.WriteLine("✅ recaptcha ng-valid — ĐÚNG");
                            return true;
                        }
                    }
                }
                catch { }

                // ═══ 5. class success ═══
                try
                {
                    var successBar = driver.FindElements(By.CssSelector(
                        "div.slider-bar.success, div.captcha-container.success, .captcha-success"));
                    if (successBar.Count > 0)
                    {
                        foreach (var el in successBar)
                        {
                            if (el.Displayed)
                            {
                                Console.WriteLine("✅ Có class success — ĐÚNG");
                                return true;
                            }
                        }
                    }
                }
                catch { }

                // ═══ 6. FALLBACK: Mặc định SAI để retry ═══
                Console.WriteLine("⚠️ Không xác định rõ — mặc định SAI để retry");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi CheckCaptchaSuccess: {ex.Message}");
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
            // ═══ 1. VNPT: ưu tiên MCCQT ═══
            var mccqt = FindElementValue(doc, ns, "MCCQT");
            if (!string.IsNullOrWhiteSpace(mccqt))
                return mccqt.Trim();

            // ═══ 2. KeySearch (một số hệ thống dùng) ═══
            var keySearch = FindElementValue(doc, ns, "KeySearch");
            if (!string.IsNullOrWhiteSpace(keySearch))
                return keySearch.Trim();

            // ═══ 3. SecretCode (một số hệ thống dùng) ═══
            var secretCode = FindElementValue(doc, ns, "SecretCode");
            if (!string.IsNullOrWhiteSpace(secretCode))
                return secretCode.Trim();

            // ═══ 4. Fkey — nằm trong <TTKhac><TTin> với <TTruong>Fkey</TTruong> ═══
            var fkey = FindTThKhacValue(doc, ns, "Fkey");
            if (!string.IsNullOrWhiteSpace(fkey))
                return fkey.Trim();

            // ═══ 5. Mã bảo mật (fallback cũ) ═══
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
        private static string FindTThKhacValue(XDocument doc, XNamespace ns, string fieldName)
        {
            try
            {
                // Tìm tất cả thẻ <TTin> ở mọi cấp
                var ttins = doc.Descendants(ns + "TTin");
                if (!ttins.Any())
                    ttins = doc.Descendants("TTin");

                foreach (var ttin in ttins)
                {
                    var truong = ttin.Element(ns + "TTruong")?.Value
                              ?? ttin.Element("TTruong")?.Value;

                    if (string.Equals(truong?.Trim(), fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        var dlieu = ttin.Element(ns + "DLieu")?.Value
                                 ?? ttin.Element("DLieu")?.Value;

                        if (!string.IsNullOrWhiteSpace(dlieu))
                            return dlieu.Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Lỗi FindTThKhacValue({fieldName}): {ex.Message}");
            }

            return null;
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
        private void KillOldOcrService()
        {
            try
            {
                var processes = Process.GetProcessesByName("ocr_service");
                if (processes.Length == 0)
                {
                    Log("ℹ️ Không có ocr_service cũ");
                    return;
                }

                Log($"⚠️ Phát hiện {processes.Length} ocr_service cũ — kill hết...");

                foreach (var p in processes)
                {
                    try
                    {
                        int pid = p.Id;
                        p.Kill();
                        p.WaitForExit(3000);
                        Log($"  ✅ Đã kill PID {pid}");
                    }
                    catch (Exception ex)
                    {
                        Log($"  ⚠️ Không kill được PID {p.Id}: {ex.Message}");
                    }
                }

                // Đợi 500ms cho OS dọn process
                System.Threading.Thread.Sleep(500);

                // Xóa file lock cũ nếu có
                try
                {
                    string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PythonOCR");
                    string requestFile = Path.Combine(baseDir, "request.txt");
                    string resultFile = Path.Combine(baseDir, "result.txt");
                    string readyFile = Path.Combine(baseDir, "ready.flag");

                    if (File.Exists(requestFile)) File.Delete(requestFile);
                    if (File.Exists(resultFile)) File.Delete(resultFile);
                    if (File.Exists(readyFile)) File.Delete(readyFile);

                    Log("  ✅ Đã xóa file lock cũ");
                }
                catch { }

                Log("✅ Đã dọn sạch OCR service cũ");
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi KillOldOcrService: {ex.Message}");
            }
        }
        private async void btnRun_Click(object sender, EventArgs e)
        {
            KillOldOcrService();
            //await TraCuuNCC(true, "3500405761", "AOEAN3B6ZH3B1FS", "13047", "D:\\DA3\\DA3\\Hoadon", 1);
            // await GetHoaDonMisa(true, "MVF0CBREBX2Q", "3502243033", "31768",1, "D:\\DA3\\DA3");  
            //await TaiBKAV(true, "3500898348", "189", "C26TSV", "b486f3cf-8754-407a-a8ed-f5206e94e7a2", @"D:\\DA3\\DA3",1);
            // await TaiSoftdream(true, "3502475274", "14933", "", "f5c5537e-b3b1-43a9-9f55-e6da4dda5ca5", "D:\\DA3\\DA3\\Hoadon", 1);
            //await TaiVisnam(true, "051080000486", "35", "", "2H0KBEXT13Q4", "D:\\DA3\\DA3\\Hoadon", 1);
            //await TaiWinTech(true, "0312643768", "2219", "C26TTC", "E2610012781966", "D:\\DA3\\DA3\\Hoadon", 1); 
            //await TaiCyber(true, "0317962038", "863", "C26TDE", "1Y7134G9ZN80", "D:\\DA3\\DA3\\Hoadon", 1);
            //TaiHDVNPTPETROLIMEX
            await TaiHDVNPTPETROLIMEX(true, "0300555450-009", "567001", "K26TXT", "C4XV4QNCG*", "D:\\DA3\\DA3\\Hoadon", 1);
            return;
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
            //try
            //{
            //    _pdfCts?.Cancel();
            //    _pdfSignal?.Release(1000);
            //    // Chờ tối đa 10s để worker dừng
            //    _pdfWorkerTask?.Wait(10000);
            //}
            //catch { }
            //base.OnFormClosing(e);
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