using BrotliSharpLib;
using ClosedXML.Excel;
using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.Data.Utils;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraPrinting.Native.Navigation;
using DevExpress.XtraPrinting.Native.WebClientUIControl;
using DevExpress.XtraWaitForm;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Wordprocessing;
using FuzzySharp;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
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
using static DevExpress.Data.Helpers.FindSearchRichParser;
//using static DevExpress.Data.Mask.Internal.MaskSettings<T>;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static ToolTaiHD.Form1;
using Process = System.Diagnostics.Process;

namespace ToolTaiHD
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
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
        #endregion

        #region Constructor & Form Load
        public bool isAutoMode { get; set; }
        private NotifyIcon notifyIcon;
        private ContextMenuStrip contextMenu;
        public Form1(bool isAutoMode = false)

        {
            InitializeComponent();
            this.isAutoMode = isAutoMode; // ✅ Lưu lại

            if (isAutoMode)
            {
                this.ShowInTaskbar = true;

                Text = "Ứng dụng - Chạy từ Task Scheduler";
                InitializeNotifyIcon();
                //RunScheduledTasks();
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

                                // ✅ Đặt con trỏ xuống cuối và scroll
                                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                                richTextBox1.ScrollToCaret();
                            }
                        }));
                    }
                    else
                    {
                        richTextBox1.Text += $"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}";

                        // ✅ Đặt con trỏ xuống cuối và scroll
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
        private void InitializeNotifyIcon()
        {
            // Tạo NotifyIcon
            notifyIcon = new NotifyIcon();

            // Set icon (bạn cần có file .ico hoặc dùng icon mặc định)
            notifyIcon.Icon = SystemIcons.Application; // Hoặc load từ file: new Icon("app.ico")

            // Set tooltip khi hover
            string appPaths = Assembly.GetExecutingAssembly().Location;

            // Lấy thư mục chứa ứng dụng
            string directoryPath = Path.GetDirectoryName(appPaths);

            // Xóa phần \bin\Debug để lấy đường dẫn gốc
            string rootDirectory = Path.GetFullPath(Path.Combine(directoryPath, @"..\.."));

            try
            { 
                notifyIcon.Text = "Tool tải tự động";
            }
            catch (Exception ex)
            {
             
            }

            string appPath = Application.StartupPath;

            // Đi lên 2 cấp để đến thư mục Resources trong source
            // Ví dụ: bin\Debug\ → ..\..\Resources\favicon.ico
            string iconPath = Path.Combine(appPath, @"..\..\Resources\favicon.ico");
            // Kiểm tra file tồn tại
            if (File.Exists(iconPath))
            {
                notifyIcon.Icon = new Icon(iconPath);
            }
            else
            {
                notifyIcon.Icon = SystemIcons.Application;
            }
            // Tạo menu chuột phải
            contextMenu = new ContextMenuStrip();

            // Thêm các menu item
            contextMenu.Items.Add("Hiện ứng dụng", null, ShowApp_Click);
            contextMenu.Items.Add("-"); // Separator
            contextMenu.Items.Add("Thoát", null, ExitApp_Click);

            // Gán menu cho NotifyIcon
            notifyIcon.ContextMenuStrip = contextMenu;

            // Xử lý sự kiện click đúp chuột vào icon
            notifyIcon.DoubleClick += NotifyIcon_DoubleClick;

            // Hiển thị NotifyIcon
            notifyIcon.Visible = true;

            // Khi form load, ẩn form (nếu cần) 
        }
        // ✅ Hàm ẩn xuống System Tray
        private void HideToSystemTray()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    this.WindowState = FormWindowState.Minimized;
                    this.ShowInTaskbar = false;
                    this.Visible = false;
                    notifyIcon.Visible = true;

                    // Hiển thị thông báo
                    notifyIcon.ShowBalloonTip(3000, "Tool tải hóa đơn",
                        "Ứng dụng đang chạy ngầm trong System Tray!",
                        ToolTipIcon.Info);
                }));
            }
            else
            {
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                this.Visible = false;
                notifyIcon.Visible = true;

                // Hiển thị thông báo
                notifyIcon.ShowBalloonTip(3000, "Tool tải hóa đơn",
                    "Ứng dụng đang chạy ngầm trong System Tray!",
                    ToolTipIcon.Info);
            }
        }
        private void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            // Double click vào icon → hiện form
            ShowApp();
        }

        private void ShowApp_Click(object sender, EventArgs e)
        {
            ShowApp();
        }

        private void ShowApp()
        {
            // Hiện form
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true; // Hiện lại trên taskbar
            this.BringToFront();

            // Focus vào form
            this.Activate();
        }

        private void ExitApp_Click(object sender, EventArgs e)
        {
            // Thoát ứng dụng
            notifyIcon.Visible = false;
            notifyIcon.Dispose();
            Application.Exit();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

            // Khi form bị minimize, ẩn đi và chỉ hiển thị icon ở system tray
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.ShowInTaskbar = false;
            }
        }
        private bool _isExiting = false; // Thêm biến cờ

        // Dọn dẹp khi form đóng
        private void SetupGridFolderColumn()
        {
            // ✅ Tạo RepositoryItemButtonEdit cho cột Folder
            RepositoryItemButtonEdit folderButton = new RepositoryItemButtonEdit();

            // ✅ Thêm icon folder (dùng image từ Resources hoặc từ file)
            folderButton.Buttons[0].Kind = DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph;
            folderButton.Buttons[0].ImageOptions.Image = Properties.Resources.open2_16x16; // Từ Resources
                                                                                           // Hoặc load từ file:
                                                                                           // folderButton.Buttons[0].ImageOptions.Image = Image.FromFile("folder.png");

            // ✅ Set tooltip
            folderButton.Buttons[0].ToolTip = "Mở thư mục";

            // ✅ Gán cho cột
            gridView1.Columns["Folder"].ColumnEdit = folderButton;

            // ✅ Bắt sự kiện click button
            folderButton.ButtonClick += FolderButton_ButtonClick;
        }

        // ✅ Sự kiện click vào button
        private void FolderButton_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            // Lấy dòng hiện tại
            int rowHandle = gridView1.FocusedRowHandle;
            if (rowHandle < 0) return;

            // Lấy giá trị cột cần thiết (ví dụ: đường dẫn)
            string folderPath = gridView1.GetRowCellValue(rowHandle, "FolderPath")?.ToString();
            // Hoặc lấy DataRow
            DataRow row = gridView1.GetDataRow(rowHandle);
            string folderPath2 = row["FolderPath"]?.ToString();

            // Mở thư mục
            if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
            {
                Process.Start("explorer.exe", folderPath);
            }
            else
            {
                XtraMessageBox.Show("Thư mục không tồn tại!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private async void Form1_Load(object sender, EventArgs e)
        {
           string computerName = Environment.MachineName;
            this.Text = $"{computerName} - Saoviet auto";
            await Task.Run(() => WaitForInternetConnection());

            radioButton1.Checked = true;
            radioButton1.Text = "Có kết nối Internet!";

            string dbPath = Path.Combine("\\\\192.168.1.90\\Ke toan 2025 New\\1 Copi vao dung 1\\Hoadon", "Tooldb.accdb");
            connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";

            string query = @"SELECT * FROM tbCompany WHERE Saoviet = ?  order by STT";
            tbCompany = ExecuteQuery(query, new OleDbParameter("?", computerName));
           
            gridControl1.DataSource = tbCompany;
            SetupGridCheckBox();
            SetupGridFolderColumn();
            // Load dữ liệu cache 


            string qrst = @"SELECT * FROM tbsetting WHERE Saoviet = ?";
            DataTable tbSetting = ExecuteQuery(qrst, new OleDbParameter("?", computerName));
            if (tbSetting.Rows.Count > 0)
            {
                int autotai = tbSetting.Rows[0].Field<int?>("AutoTai") ?? 0;
                string Block1 = tbSetting.Rows[0].Field<string>("Block1");
                string Block2 = tbSetting.Rows[0].Field<string>("Block2");
                string Block3 = tbSetting.Rows[0].Field<string>("Block3");

                checkEdit1.Checked = autotai == 1 ? true : false;
                chkMoc1.Checked = !string.IsNullOrEmpty(Block1) ? true : false;
                chkMoc2.Checked = !string.IsNullOrEmpty(Block2) ? true : false;
                chkMoc3.Checked = !string.IsNullOrEmpty(Block3) ? true : false;
                txtSoluongtai.Text = tbSetting.Rows[0].Field<int>("Soluongtai").ToString();
                txttimeout.Text = tbSetting.Rows[0].Field<int>("Timeout").ToString();
                if (chkMoc1.Checked)
                {
                    txtBlock1.Enabled = true;
                    txtBlock1.Text = Block1;
                }
                if (chkMoc2.Checked)
                {
                    txtBlock2.Enabled = true;
                    txtBlock2.Text = Block1;
                }
                if (chkMoc3.Checked)
                {
                    txtBlock3.Enabled = true;
                    txtBlock3.Text = Block1;
                }
            }
           

            if (isAutoMode)
            {
                btnRun.PerformClick();
                HideToSystemTray();
            } 
        }
        #endregion

        #region Internet Connection Check
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
                        "Cảnh báo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );

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
            {
                radioButton1.Invoke(new Action(() =>
                {
                    radioButton1.Text = text;
                    radioButton1.Checked = isChecked;
                }));
            }
            else
            {
                radioButton1.Text = text;
                radioButton1.Checked = isChecked;
            }
        }

        public bool IsInternetAvailable()
        {
            try
            {
                using (var ping = new Ping())
                {
                    var reply = ping.Send("8.8.8.8", 3000);
                    if (reply != null && reply.Status == IPStatus.Success)
                    {
                        return true;
                    }
                }

                using (var ping = new Ping())
                {
                    var reply = ping.Send("1.1.1.1", 3000);
                    return reply != null && reply.Status == IPStatus.Success;
                }
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region Task Scheduler Check
        private bool IsRunningFromTaskScheduler()
        {
            if (Environment.UserName.Equals("SYSTEM", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            try
            {
                Process currentProcess = Process.GetCurrentProcess();
                Process parentProcess = GetParentProcess(currentProcess.Id);

                if (parentProcess != null)
                {
                    string parentName = parentProcess.ProcessName.ToLower();

                    if (parentName.Contains("svchost") ||
                        parentName.Contains("taskeng") ||
                        parentName.Contains("taskschd") ||
                        parentName == "taskhostw" ||
                        parentName == "taskhostex")
                    {
                        parentProcess.Dispose();
                        return true;
                    }

                    parentProcess.Dispose();
                }
            }
            catch { }

            uint sessionId = WTSGetActiveConsoleSessionId();
            if (sessionId == 0 || sessionId == 0xFFFFFFFF)
            {
                return true;
            }

            string taskName = Environment.GetEnvironmentVariable("SCHEDULED_TASK_NAME");
            string taskFolder = Environment.GetEnvironmentVariable("SCHEDULED_TASK_FOLDER");

            if (!string.IsNullOrEmpty(taskName) || !string.IsNullOrEmpty(taskFolder))
            {
                return true;
            }

            return false;
        }

        private Process GetParentProcess(int id)
        {
            try
            {
                Process process = Process.GetProcessById(id);
                PerformanceCounter pc = new PerformanceCounter("Process", "Creating Process ID", process.ProcessName);
                int parentId = (int)pc.RawValue;

                if (parentId > 0)
                {
                    return Process.GetProcessById(parentId);
                }
            }
            catch { }

            return null;
        }

        private void RunScheduledTasks()
        {
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
        }

        private void ShowNormalUI()
        {
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
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
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }
                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                        {
                            dataAdapter.Fill(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Lỗi ExecuteQuery: {ex.Message}");
                }
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
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }
                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                        {
                            dataAdapter.Fill(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Lỗi ExecuteQuery2: {ex.Message}");
                }
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
                    if (parameters != null)
                        command.Parameters.AddRange(parameters);
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
                    if (parameters != null)
                        command.Parameters.AddRange(parameters);
                    command.ExecuteNonQuery();
                }

                using (OleDbCommand idCommand = new OleDbCommand("SELECT @@IDENTITY", connection))
                {
                    object result = idCommand.ExecuteScalar();
                    return Convert.ToInt32(result);
                }
            }
        }
        #endregion

        #region Load Cache Data
        private void LoadHoadonCT(string connectionst)
        {
            try
            {
                string query = @"
                    SELECT 
                        hd.SoHD,
                        hd.KyHieu,
                        hd.NgayPH,
                        hd.MaKhachHang,
                        ct.NgayCT,
                        ct.MaLoai
                    FROM 
                        Hoadon hd
                    INNER JOIN 
                        Chungtu ct ON hd.MaSo = ct.MaSo
                    WHERE 
                        hd.KyHieu <> '...'";

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
            catch (Exception ex)
            {
                Log($"Lỗi LoadHoadonCT: {ex.Message}");
            }
        }
        DataTable tbimports;
        private void Loadtbimport(string conectionst)
        {
            try
            {
                string query = "SELECT * FROM tbimport";
                tbimports = ExecuteQuery2(query, conectionst);
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
            catch (Exception ex)
            {
                Log($"Lỗi Loadtbimport: {ex.Message}");
            }
        }
        #endregion

        #region Token & Authentication
        public void Gettokken(string username, string password, ref string currentToken, string connectist)
        {
            int maxRetry = 3;
            int retryCount = 0;

            while (retryCount < maxRetry)
            {
                retryCount++;
                try
                {
                    using (var client = new HttpClient())
                    {
                        HttpResponseMessage response = null;
                        string url = "https://hoadondientu.gdt.gov.vn/api/captcha";
                        int captchaRetry = 0;
                        int maxCaptchaRetry = 10;

                        while (captchaRetry < maxCaptchaRetry)
                        {
                            try
                            {
                                response = client.GetAsync(url).Result;
                                if (response.IsSuccessStatusCode)
                                {
                                    byte[] captchaBytes = response.Content.ReadAsByteArrayAsync().Result;
                                    if (captchaBytes.Length > 0)
                                    {
                                        break;
                                    }
                                }
                            }
                            catch { }

                            captchaRetry++;
                            Thread.Sleep(1000);
                        }

                        if (response == null || !response.IsSuccessStatusCode)
                        {
                            currentToken = "";
                            return;
                        }

                        string responseBody = response.Content.ReadAsStringAsync().Result;
                        MyJson myJson = JsonConvert.DeserializeObject<MyJson>(responseBody);
                        string filePath = AppDomain.CurrentDomain.BaseDirectory + "output.svg";
                        File.WriteAllText(filePath, myJson.Content);
                        Thread.Sleep(50);

                        SvgCaptchaSolver solver = new SvgCaptchaSolver();
                        string captchaResult = solver.SolveCaptcha(filePath);

                        url = "https://hoadondientu.gdt.gov.vn/api/security-taxpayer/authenticate";
                        var payload = new
                        {
                            username = username,
                            password = password,
                            cvalue = captchaResult,
                            ckey = myJson.Key
                        };

                        string json = JsonConvert.SerializeObject(payload);
                        var content = new StringContent(json, Encoding.UTF8, "application/json");
                        response = client.PostAsync(url, content).Result;

                        if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                        {
                            Log($"⚠️ Đăng nhập thất bại lần {retryCount}");
                           
                            Thread.Sleep(2000);
                            continue;
                        }

                        response.EnsureSuccessStatusCode();
                        Thread.Sleep(50);
                        responseBody = response.Content.ReadAsStringAsync().Result;
                        var tokenResponse = JsonConvert.DeserializeObject<TokenResponse>(responseBody);

                        currentToken = tokenResponse.token;
                        this.tokken = tokenResponse.token;

                        Log($"✅ Token trả về: {currentToken.Substring(0, Math.Min(20, currentToken.Length))}...");
                        
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log($"⚠️ Lỗi lấy token lần {retryCount}: {ex.Message}");
                   
                    if (retryCount >= maxRetry)
                    {
                        currentToken = "";
                        return;
                    }
                    Thread.Sleep(2000);
                }
            }

            currentToken = "";
        }

        private async Task<string> GetTokenForCompanyAsync(string username, string password, string connectionString)
        {
            try
            {
                string token = await Task.Run(() =>
                {
                    string currentToken = "";
                    Gettokken(username, password, ref currentToken, connectionString);
                    return currentToken;
                });
                return token;
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi lấy token: {ex.Message}");
               
                return "";
            }
        }
        #endregion

        #region Company HttpClient
        public class CompanyHttpClient : IDisposable
        {
            private readonly HttpClient _httpClient;
            private readonly string _companyName;
            private readonly string _token;
            private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);

            public CompanyHttpClient(string token, string companyName)
            {
                _token = token;
                _companyName = companyName;

                _httpClient = new HttpClient();
                _httpClient.Timeout = TimeSpan.FromSeconds(120);
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/octet-stream"));
                _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
            }

            public async Task<HttpResponseMessage> GetAsync(string url)
            {
                await _semaphore.WaitAsync();
                try
                {
                    return await _httpClient.GetAsync(url);
                }
                finally
                {
                    _semaphore.Release();
                }
            }

            public async Task<bool> DownloadFileAsync(string url, string filePath,int timeoutSeconds)
            {
                int maxRetry = 1; 

                for (int retry = 1; retry <= maxRetry; retry++)
                {
                    try
                    {
                        using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds)))
                        {
                            var response = await GetAsync(url);

                            if ((int)response.StatusCode == 429)
                            {
                                Log($"⚠️ {_companyName}: Rate limit! Chờ 5s...");
                                await Task.Delay(5000);
                                continue;
                            }

                            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                            {
                                Log($"❌ {_companyName}: Token hết hạn!");
                                return false;
                            }

                            if (response.IsSuccessStatusCode)
                            {
                                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                                {
                                    // ❌ KHÔNG truyền token vào đây
                                    await response.Content.CopyToAsync(fs);
                                }
                                return true;
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        Log($"⏰ {_companyName}: Timeout lần {retry}/{maxRetry} sau {timeoutSeconds}s");
                    }
                    catch (Exception ex)
                    {
                        Log($"⚠️ {_companyName}: Lỗi lần {retry}: {ex.Message}");
                        if (retry < maxRetry)
                            await Task.Delay(retry * 2000);
                    }
                }
                return false;
            }

            public async Task<string> GetStringAsync(string url)
            {
                var response = await GetAsync(url);
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsStringAsync();
            }

            public void Dispose()
            {
                _httpClient?.Dispose();
                _semaphore?.Dispose();
            }
        }
        #endregion

        #region Main Processing
        //private async void btnRun_Click(object sender, EventArgs e)
        //{
        //    btnRun.Enabled = false;

        //    try
        //    {
        //        string query = @"SELECT * FROM tbCompany WHERE IsRun = 1";
        //        tbCompany = ExecuteQuery(query);

        //        if (tbCompany.Rows.Count == 0)
        //        {
        //            MessageBox.Show("Không có công ty nào đang hoạt động!");
        //            return;
        //        }

        //        Log($"🚀 Bắt đầu xử lý {tbCompany.Rows.Count} công ty (TUẦN TỰ)");
        //        foreach (DataRow item in tbCompany.Rows)
        //        {
        //            string vbdbpath = item["Dbpath"]?.ToString() ?? "";
        //            string companyName = item["Name"]?.ToString() ?? "Unknown";

        //            if (string.IsNullOrEmpty(vbdbpath))
        //            {
        //                Log($"⚠️ {companyName}: Không có Dbpath, bỏ qua!");
        //                continue;
        //            }

        //            await TaihoadonCongty(vbdbpath, item);

        //            Log($"⏳ Chờ 5s trước khi xử lý công ty tiếp theo...");
        //            await Task.Delay(5000);
        //        }

        //        Log($"✅ Hoàn thành xử lý tất cả công ty!");
        //    }
        //    catch (Exception ex)
        //    {
        //        Log($"❌ Lỗi: {ex.Message}");
        //    }
        //    finally
        //    {
        //        btnRun.Enabled = true;
        //    }
        //}
       

        // ✅ Hàm cập nhật status an toàn trên UI thread
        private void UpdateStatusOnUI(DataRow row, string statusText)
        {
            try
            {
                if (row == null) return;

                // ✅ Kiểm tra nếu đang ở UI thread hay không
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() =>
                    {
                        try
                        {
                            if (!row.Table.Columns.Contains("Status"))
                            {
                                // Nếu chưa có cột Status, thêm vào
                                row.Table.Columns.Add("Status", typeof(string));
                            }
                            row["Status"] = statusText;
                            gridControl1.DataSource = tbCompany;
                        }
                        catch (Exception ex)
                        {
                            Log($"⚠️ Lỗi update status: {ex.Message}");
                        }
                    }));
                }
                else
                {
                    if (!row.Table.Columns.Contains("Status"))
                    {
                        row.Table.Columns.Add("Status", typeof(string));
                    }
                    row["Status"] = statusText;
                    gridControl1.DataSource = tbCompany;
                }
            }
            catch (Exception ex)
            {
                Log($"⚠️ Lỗi UpdateStatusOnUI: {ex.Message}");
            }
        }
        private async Task TaihoadonCongty(string vbdbpath, DataRow dtrow)
        {
            if (string.IsNullOrEmpty(vbdbpath))
            {
                Log($"⚠️ Dbpath trống cho {dtrow["Name"]}");
                return;
            }

            string connectionString2 = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                       "Data Source=" + vbdbpath + ";" +
                                       "Jet OLEDB:Database Password=1@35^7*9)1;";
            string qrkh = "SELECT * FROM KhachHang"; // Giả sử bạn muốn lấy tất cả dữ liệu từ bảng KhachHang
            tbKhachhang = ExecuteQuery2(qrkh, connectionString2);
         
            try
            {
                string query = @"SELECT * FROM tbRegister";
                var getResgistry = ExecuteQuery2(query, connectionString2);
               string mstcongtys = getResgistry.Rows[0]["Username"].ToString();
                if (getResgistry.Rows.Count == 0)
                {
                    Log($"⚠️ Không tìm thấy tbRegister cho {dtrow["Name"]}");
                    return;
                }

                string companyName = dtrow["Name"].ToString();
                string username = getResgistry.Rows[0]["Username"].ToString();
                string password = getResgistry.Rows[0]["Password"].ToString();
                string savedPath = getResgistry.Rows[0]["Hoadonpath"].ToString();

                Log($"[{companyName}] ▶️ Bắt đầu xử lý...");

                string token = await GetTokenForCompanyAsync(username, password, connectionString2);

                if (string.IsNullOrEmpty(token))
                {
                    Log($"❌ {companyName}: Đăng nhập thất bại!");
                    return;
                }

                using (var companyClient = new CompanyHttpClient(token, companyName))
                {
                    // ============================================================
                    // ✅ BƯỚC 1: TẢI EXCEL ĐẦU VÀO (HDVao)
                    // ============================================================
                    Log($"[{companyName}] 📥 Bắt đầu tải Excel đầu vào...");

                    bool t1 = await XulyexelvaoAsync(companyClient, 1, savedPath, companyName, mstcongtys);
                    bool t2 = await XulyexelvaoAsync(companyClient, 2, savedPath, companyName, mstcongtys);
                    bool t3 = await XulyexelvaoAsync(companyClient, 3, savedPath, companyName, mstcongtys);

                    if (t1 && t2 && t3)
                    {
                        Log($"✅ {companyName}: Tải Excel đầu vào thành công!");
                    }
                    else
                    {
                        Log($"⚠️ {companyName}: Có lỗi khi tải Excel đầu vào!");
                    }

                    // ============================================================
                    // ✅ BƯỚC 2: TẢI HÓA ĐƠN ĐẦU VÀO
                    // ============================================================
                    Log($"[{companyName}] 📥 Bắt đầu tải hóa đơn đầu vào...");

                    string currentYear = $"HD{DateTime.Now.Year}";
                    string directoryPathVao = Path.Combine(savedPath, currentYear, "HDVao");

                    var invoicesVao = await GetListHoaDonCanTai(username, savedPath, 1);
                    if (invoicesVao.Count > 0)
                    {
                      await TaiHangLoatHoaDon(companyClient, invoicesVao, "đầu vào", companyName, dtrow);
                    }
                    else
                    {
                        Log($"📭 {companyName}: Không có hóa đơn đầu vào mới");
                    }

                    // ============================================================
                    // ✅ BƯỚC 3: TẢI EXCEL ĐẦU RA (HDRa)
                    // ============================================================
                    Log($"[{companyName}] 📤 Bắt đầu tải Excel đầu ra...");

                    bool t1r = await XulyexelraAsync(companyClient, 1, savedPath, companyName, mstcongtys);
                    bool t2r = await XulyexelraAsync(companyClient, 2, savedPath, companyName, mstcongtys);

                    if (t1r && t2r)
                    {
                        Log($"✅ {companyName}: Tải Excel đầu ra thành công!");
                    }
                    else
                    {
                        Log($"⚠️ {companyName}: Có lỗi khi tải Excel đầu ra!");
                    }


                    // ============================================================
                    // ✅ BƯỚC 4: TẢI HÓA ĐƠN ĐẦU RA
                    // ============================================================
                    Log($"[{companyName}] 📤 Bắt đầu tải hóa đơn đầu ra...");

                    string directoryPathRa = Path.Combine(savedPath, currentYear, "HDRa");
                    string qriv = @"SELECT * FROM tbInvoiceInfo";
                    var tbInvoiceInfo = ExecuteQuery2(qriv, connectionString2);
                    var invoicesRa = await GetListHoaDonCanTai(username, savedPath, 2);
                    if (invoicesRa.Count > 0 || tbInvoiceInfo.Rows.Count > 0)
                    {
                        //Kiểm tra xem có đăng ký invoice không
                     
                        if(tbInvoiceInfo.Rows.Count > 0)
                        {
                            Tainhacungcap(vbdbpath);
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
                //Log($"❌ {companyName}: Lỗi TaihoadonCongty: {ex.Message}");
            }
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
        public static LoginResponse loginResponse { get; set; } = new LoginResponse();
        public static UseCookie useCookie { get; set; } = new UseCookie();
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
            // Có thể thêm các trường khác nếu cần sử dụng
        }
        public class SearchResponse
        {
            public int code { get; set; }
            public string message { get; set; }
            public SearchData data { get; set; }
        }
        private async void Tainhacungcap(string vbdbpath)
        {
            string connectionString2 = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                      "Data Source=" + vbdbpath + ";" +
                                      "Jet OLEDB:Database Password=1@35^7*9)1;";
            string qrq = "SELECT * FROM tbInvoiceInfo";
            var dtInvoiceInfo = ExecuteQuery2(qrq, connectionString2);
            var row = dtInvoiceInfo.Rows[0];

            string username = row["Username"]?.ToString();
            string password = row["Password"]?.ToString();

            var url = "https://vinvoice.viettel.vn/api/auth/login";
            using (HttpClientHandler handler = new HttpClientHandler())
            {
                // Tùy chọn: tự động xử lý cookie
                handler.UseCookies = true;
                handler.CookieContainer = new CookieContainer();

                using (HttpClient client = new HttpClient(handler))
                {
                    // giống Postman
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                    client.DefaultRequestHeaders.Add("Accept", "application/json");

                    var json = $@"{{
                        ""username"": ""{username}"",
                        ""password"": ""{password}"",
                        ""rememberMe"": false,
                        ""captcha"": """"
                    }}"; 
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(url, content);
                    if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                    { 
                        Application.DoEvents();
                        Thread.Sleep(3000);
                        Application.Exit();
                    }
                    // *** CÁCH LẤY COOKIE ***
                    // Lấy tất cả cookies từ response
                    var cookies = handler.CookieContainer.GetCookies(new Uri(url));

                    foreach (System.Net.Cookie cookie in cookies)
                    {
                        //MessageBox.Show($"Cookie: {cookie.Name} = {cookie.Value}");

                        // Nếu bạn muốn lấy riêng cookie __cf_bm
                        if (cookie.Name == "__cf_bm")
                        {
                            string cf_bm_value = cookie.Value;
                            useCookie.__cf_bm = cf_bm_value;
                        }
                        if (cookie.Name == "JSESSIONID")
                        {
                            string JSESSIONID_value = cookie.Value;
                            useCookie.JSESSIONID = JSESSIONID_value;
                        }
                        if (cookie.Name == "access_token")
                        {
                            string access_token_value = cookie.Value;
                            useCookie.access_token = access_token_value;
                        }
                        if (cookie.Name == "session_token")
                        {
                            string session_token_value = cookie.Value;
                            useCookie.session_token = session_token_value;
                        }
                    }

                    var result = await response.Content.ReadAsStringAsync();
                    loginResponse = JsonConvert.DeserializeObject<LoginResponse>(result);
                    if (loginResponse != null)
                    { 
                        Application.DoEvents();

                        try
                        {
                            // ========================================
                            // GỌI API LẤY DANH SÁCH HÓA ĐƠN
                            // ========================================
                            int page = 0;
                            int size = 1000;
                            int supplierId = 103807;

                            // Lấy ngày đầu tháng
                            DateTime firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                            // Lấy ngày cuối tháng
                            DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

                            // Format theo yêu cầu
                            string fromDate = firstDayOfMonth.AddHours(-7).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
                            string toDate = firstDayOfMonth.AddDays(1).AddSeconds(-1).AddHours(-7).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

                            string urls = $"https://vinvoice.viettel.vn/api/cluster7/services/einvoiceapplication/api/invoice/search" +
                                         $"?page={page}" +
                                         $"&size={size}" +
                                         $"&createdDate.greaterThanOrEqual={Uri.EscapeDataString(fromDate)}" +
                                         $"&createdDate.lessThanOrEqual={Uri.EscapeDataString(toDate)}" +
                                         $"&supplierId.equals={supplierId}" +
                                         $"&dateType.equals=0" +
                                         $"&invoiceStatus.equals=1" +
                                         $"&invoiceTypeId.notEquals=52" +
                                         $"&sort=issueDate,desc" +
                                         $"&sort=invoiceNumber,desc";
                            // Thêm Authorization header
                            if (!string.IsNullOrEmpty(useCookie.access_token))
                            {
                                client.DefaultRequestHeaders.Authorization =
                                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", useCookie.access_token);
                            }

                            var responses = await client.GetAsync(urls);

                            if (responses.IsSuccessStatusCode)
                            {
                                string jsonResult = await responses.Content.ReadAsStringAsync();
                                var searchResult = JsonConvert.DeserializeObject<SearchResponse>(jsonResult);

                                // ========================================
                                // KIỂM TRA VÀ XỬ LÝ DANH SÁCH HÓA ĐƠN
                                // ========================================
                                if (searchResult != null && searchResult.code == 200 && searchResult.data != null)
                                {
                                    var invoices = searchResult.data.content; 
                                    Application.DoEvents();

                                    // Tạo thư mục lưu file
                                    //string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Invoices");
                                    //Directory.CreateDirectory(folderPath);

                                    int successCount = 0;
                                    int total = invoices.Count;
                                    string pathravao = "HDRa";
                                    string pathYear = $"HD{firstDayOfMonth.Year}";
                                    string query = @"SELECT * FROM tbRegister";
                                    var getResgistry = ExecuteQuery2(query, connectionString2);
                                    string savedPath = getResgistry.Rows[0]["Hoadonpath"].ToString();
                                    for (int i = 0; i < invoices.Count; i++)
                                    {
                                        var invoice = invoices[i];
                                        string filename = $"{invoice.createdDate.ToString("yyyyMMdd")}_{mstcongty}_{invoice.invoiceNumber}_{invoice.invoiceSeri}.xml";
                                        string filepath = Path.Combine(savedPath, pathYear, pathravao, firstDayOfMonth.Month.ToString(), filename);
                                        string filenamepdf = $"{invoice.createdDate.ToString("yyyyMMdd")}_{mstcongty}_{invoice.invoiceNumber}_{invoice.invoiceSeri}.pdf";
                                        string filepathpdf = Path.Combine(savedPath, pathYear, pathravao, firstDayOfMonth.Month.ToString(), filenamepdf);

                                        DownloadInvoiceXml(client, invoice.id.ToString(), filepath);
                                        DownloadInvoicePdf(client, invoice.id.ToString(), filepathpdf); 
                                        Application.DoEvents();
                                        // Chờ 500ms để tránh rate limit
                                        await Task.Delay(150);
                                    }  
                                    //MessageBox.Show($"Hoàn thành tải {successCount}/{total} hóa đơn!\nLưu tại: {filepath}",
                                    //               "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                { 
                                    MessageBox.Show("Không tìm thấy hóa đơn nào trong khoảng thời gian này!",
                                                   "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                            else
                            {
                                string error = await responses.Content.ReadAsStringAsync(); 
                                MessageBox.Show($"Lỗi tìm kiếm: {responses.StatusCode}\n{error}",
                                               "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        catch (Exception ex)
                        { 
                            MessageBox.Show($"Lỗi xử lý: {ex.Message}",
                                           "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }
        private async Task DownloadInvoicePdf(HttpClient client, string invoiceId, string filePath)
        {
            try
            {
                // URL API tạo PDF
                string url = $"https://vinvoice.viettel.vn/api/cluster7/services/einvoicequery/api/invoice/gen-pdf?id={invoiceId}";
                 
                Application.DoEvents();

                // Thêm header Authorization nếu cần
                if (!string.IsNullOrEmpty(useCookie.access_token))
                {
                    client.DefaultRequestHeaders.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", useCookie.access_token);
                }

                // Header cho phản hồi dạng PDF
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/pdf"));

                var response = await client.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    // Đọc dữ liệu dạng byte (PDF là binary)
                    byte[] fileBytes = await response.Content.ReadAsByteArrayAsync();

                    // Lưu file PDF
                    //string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Invoices");
                    //string filePath = Path.Combine(folderPath, $"Invoice_{invoiceId}.pdf");

                    File.WriteAllBytes(filePath, fileBytes);
                     
                    //MessageBox.Show($"Tải file PDF thành công!\nLưu tại: {filePath}",
                    //               "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync(); ;
                    //MessageBox.Show($"Lỗi tải PDF: {response.StatusCode}\n{error}",
                    //               "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            { 
                //MessageBox.Show($"Lỗi tải PDF: {ex.Message}",
                //               "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private async Task DownloadInvoiceXml(HttpClient client, string invoiceId, string filePath)
        {
            try
            {
                string url = $"https://vinvoice.viettel.vn/api/cluster7/services/einvoiceapplication/api/invoice/downloadInvoiceFileXmlById?invoiceId={invoiceId}";
                 
                Application.DoEvents();

                // Thêm header Authorization nếu cần
                if (!string.IsNullOrEmpty(useCookie.access_token))
                {
                    client.DefaultRequestHeaders.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", useCookie.access_token);
                }

                // Thêm header để nhận XML
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/xml"));

                var response = await client.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    // Lưu file XML
                    string xmlContent = await response.Content.ReadAsStringAsync();

                    // Lưu vào file
                    //string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Invoices");
                    //string filePath = Path.Combine(folderPath, $"Invoice_{invoiceId}.xml");

                    File.WriteAllText(filePath, xmlContent);
                     
                    //MessageBox.Show($"Tải file XML thành công!\nLưu tại: {filePath}",
                    //               "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Nếu muốn lưu dưới dạng file .zip (nếu API trả về zip)
                    // byte[] fileBytes = await response.Content.ReadAsByteArrayAsync();
                    // File.WriteAllBytes(filePath.Replace(".xml", ".zip"), fileBytes);
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync(); 
                    //MessageBox.Show($"Lỗi tải file: {response.StatusCode}\n{error}",
                    //               "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            { 
                //MessageBox.Show($"Lỗi tải file: {ex.Message}",
                //               "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private async Task TaihoadonCongtyOld(string vbdbpath, DataRow dtrow)
        {
            if (string.IsNullOrEmpty(vbdbpath))
            {
                Log($"⚠️ Dbpath trống cho {dtrow["Name"]}");
                return;
            }

            string connectionString2 = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                       "Data Source=" + vbdbpath + ";" +
                                       "Jet OLEDB:Database Password=1@35^7*9)1;";

            try
            {
                string query = @"SELECT * FROM tbRegister";
                var getResgistry = ExecuteQuery2(query, connectionString2);
                mstcongty = getResgistry.Rows[0]["Username"].ToString();
                if (getResgistry.Rows.Count == 0)
                {
                    Log($"⚠️ Không tìm thấy tbRegister cho {dtrow["Name"]}");
                    return;
                }

                string companyName = dtrow["Name"].ToString();
                string username = getResgistry.Rows[0]["Username"].ToString();
                string password = getResgistry.Rows[0]["Password"].ToString();
                string savedPath = getResgistry.Rows[0]["Hoadonpath"].ToString();

                Log($"▶️ Bắt đầu xử lý công ty: {companyName}");

                string token = await GetTokenForCompanyAsync(username, password, connectionString2);

                if (string.IsNullOrEmpty(token))
                {
                    Log($"❌ {companyName}: Đăng nhập thất bại!");
                    return;
                }

                using (var companyClient = new CompanyHttpClient(token, companyName))
                {
                    // ========== XỬ LÝ ĐẦU VÀO (HDVao) ==========
                    Log($"📥 {companyName}: Bắt đầu xử lý đầu vào...");

                    bool t1 = await XulyexelvaoAsync(companyClient, 1, savedPath, companyName, mstcongty);
                    bool t2 = await XulyexelvaoAsync(companyClient, 2, savedPath, companyName, mstcongty);
                    bool t3 = await XulyexelvaoAsync(companyClient, 3, savedPath, companyName, mstcongty);

                    if (t1 && t2 && t3)
                    {
                        Log($"✅ {companyName}: Tải Excel đầu vào thành công!");
                    }
                    else
                    {
                        Log($"⚠️ {companyName}: Có lỗi khi tải Excel đầu vào!");
                    }

                    // Tải hóa đơn đầu vào
                    string currentYear = $"HD{DateTime.Now.Year}";
                    string directoryPathVao = Path.Combine(savedPath, currentYear, "HDVao");

                    var invoicesVao = await GetListHoaDonCanTai(username, savedPath, 1);
                    if (invoicesVao.Count > 0)
                    {
                        await TaiHangLoatHoaDon(companyClient, invoicesVao, "đầu vào", companyName);
                    }
                    else
                    {
                        Log($"📭 {companyName}: Không có hóa đơn đầu vào mới");
                    }

                    // ========== XỬ LÝ ĐẦU RA (HDRa) ==========
                    Log($"📤 {companyName}: Bắt đầu xử lý đầu ra...");

                    // Đầu ra chỉ có 2 loại: Hóa đơn điện tử và Hóa đơn máy tính tiền
                    bool t1r = await XulyexelraAsync(companyClient, 1, savedPath, companyName, mstcongty); // Hóa đơn điện tử
                    bool t2r = await XulyexelraAsync(companyClient, 2, savedPath, companyName, mstcongty); // Hóa đơn máy tính tiền

                    if (t1r && t2r)
                    {
                        Log($"✅ {companyName}: Tải Excel đầu ra thành công!");
                    }
                    else
                    {
                        Log($"⚠️ {companyName}: Có lỗi khi tải Excel đầu ra!");
                    }

                    // Tải hóa đơn đầu ra
                    string directoryPathRa = Path.Combine(savedPath, currentYear, "HDRa");

                    var invoicesRa = await GetListHoaDonCanTai(username, savedPath, 2);
                    if (invoicesRa.Count > 0)
                    {
                        await TaiHangLoatHoaDon(companyClient, invoicesRa, "đầu ra", companyName);
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
                Log($"❌ Lỗi TaihoadonCongty: {ex.Message}");
            }
        }
        #endregion

        #region Excel Download - Đầu Vào
        public async Task<bool> XulyexelvaoAsync(CompanyHttpClient client, int _type, string savedPath, string companyName,string mstcongty)
        {
            try
            {
                DateTime dtFrom = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                DateTime dtTo = DateTime.Now;

                string formattedDate1 = dtFrom.ToString("dd/MM/yyyyTHH:mm:ss");
                string formattedDate2 = dtTo.ToString("dd/MM/yyyyTHH:mm:ss");

                string url, filename, fileType;

                switch (_type)
                {
                    case 1:
                        url = $"https://hoadondientu.gdt.gov.vn/api/query/invoices/export-excel-sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2};ttxly==5%20%20%20%20&type=purchase";
                        filename = $"{mstcongty}_HDDienTuDaCapMa.xlsx";
                        fileType = "hóa đơn điện tử có mã";
                        break;
                    case 2:
                        url = $"https://hoadondientu.gdt.gov.vn/api/query/invoices/export-excel-sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2};ttxly==6%20%20%20%20&type=purchase";
                        filename = $"{mstcongty}_HDDienTuKhongMa.xlsx";
                        fileType = "hóa đơn điện tử không mã";
                        break;
                    case 3:
                        url = $"https://hoadondientu.gdt.gov.vn/api/sco-query/invoices/export-excel-sold?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2};ttxly==8%20%20%20%20&type=purchase";
                        filename = $"{mstcongty}_HDDienTuMayTinhTien.xlsx";
                        fileType = "hóa đơn máy tính tiền";
                        break;
                    default:
                        return false;
                }

                string currentYear = $"HD{DateTime.Now.Year}";
                string directoryPath = Path.Combine(savedPath, currentYear, "HDVao", DateTime.Now.Month.ToString());
                Directory.CreateDirectory(directoryPath);

                string filePath = Path.Combine(directoryPath, filename);

                if (File.Exists(filePath))
                {
                    TimeSpan ts = DateTime.Now - File.GetLastWriteTime(filePath);
                    if (ts.TotalMinutes < 120)
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
                    Log($"✅ {companyName}: Đã tải xong {fileType} đầu vào");
                    return true;
                }
                else
                {
                    Log($"❌ {companyName}: Tải {fileType} đầu vào thất bại!");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"❌ {companyName}: Lỗi XulyexelvaoAsync: {ex.Message}");
                return false;
            }
        }
        #endregion

        #region Excel Download - Đầu Ra
        #region Excel Download - Đầu Ra (HDRa)
        public async Task<bool> XulyexelraAsync(CompanyHttpClient client, int _type, string savedPath, string companyName,string mstcongtys)
        {
            try
            {
                DateTime dtFrom = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                DateTime dtTo = dtFrom.AddMonths(1).AddDays(-1); // Lấy đến cuối tháng

                string formattedDate1 = dtFrom.ToString("dd/MM/yyyyTHH:mm:ss");
                string formattedDate2 = dtTo.ToString("dd/MM/yyyyTHH:mm:ss");

                string url, filename, fileType;

                switch (_type)
                {
                    case 1: // Hóa đơn điện tử đầu ra
                        url = $"https://hoadondientu.gdt.gov.vn/api/query/invoices/export-excel?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2}";
                        filename = $"{mstcongtys}_Hoadondientu.xlsx";
                        fileType = "hóa đơn điện tử (đầu ra)";
                        break;

                    case 2: // Hóa đơn máy tính tiền đầu ra
                        url = $"https://hoadondientu.gdt.gov.vn/api/sco-query/invoices/export-excel?sort=tdlap:desc,khmshdon:asc,shdon:desc&search=tdlap=ge={formattedDate1};tdlap=le={formattedDate2}";
                        filename = $"{mstcongtys}_HDDienTuMayTinhTien.xlsx";
                        fileType = "hóa đơn máy tính tiền (đầu ra)";
                        break;

                    default:
                        return false;
                }

                string currentYear = $"HD{DateTime.Now.Year}";
                string directoryPath = Path.Combine(savedPath, currentYear, "HDRa", DateTime.Now.Month.ToString());
                Directory.CreateDirectory(directoryPath);

                string filePath = Path.Combine(directoryPath, filename);

                // Kiểm tra file đã tồn tại
                if (File.Exists(filePath))
                {
                    TimeSpan ts = DateTime.Now - File.GetLastWriteTime(filePath);
                    if (ts.TotalMinutes < 120)
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
                    Log($"✅ {companyName}: Đã tải xong {fileType}");
                    return true;
                }
                else
                {
                    Log($"❌ {companyName}: Tải {fileType} thất bại!");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"❌ {companyName}: Lỗi XulyexelraAsync: {ex.Message}");
                return false;
            }
        }
        #endregion
        #endregion

        #region Get List Invoices
        private async Task<List<InvoiceInfo>> GetListHoaDonCanTaiNoneExcel(string mstcongty, string savedPath, int type,string token)
        {
            if (type == 1)
            {
                using (var client = new HttpClient())
                {
                    // Thiết lập timeout để tránh treo
                    client.Timeout = TimeSpan.FromSeconds(30);

                    // *** BƯỚC QUAN TRỌNG: Thêm token vào Header ***
                    client.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Bearer", token);

                    // Thêm header chấp nhận dữ liệu JSON
                    client.DefaultRequestHeaders.Accept.Add(
                        new MediaTypeWithQualityHeaderValue("application/json")); 
                    try
                    {
                        string apiUrl = "https://hoadondientu.gdt.gov.vn/api/query/invoices/purchase?sort=tdlap:desc&size=15&search=tdlap=ge=25/06/2026T00:00:00;tdlap=le=24/07/2026T23:59:59;ttxly==5";
                        // Gửi request GET
                        HttpResponseMessage response = await client.GetAsync(apiUrl);

                        // Đảm bảo request thành công
                        response.EnsureSuccessStatusCode();

                        // Đọc nội dung trả về dưới dạng string
                        string responseBody = await response.Content.ReadAsStringAsync();

                        // Log thành công
                        Log($"✅ Lấy danh sách hóa đơn thành công.");
                        return null;
                    }
                    catch (HttpRequestException ex)
                    {
                        // Xử lý lỗi HTTP (ví dụ: 401, 404, 500...)
                        Log($"❌ Lỗi khi gọi API: {ex.Message}");
                        return null;
                    }
                }
            }
            return null;
        }
        private async Task<List<InvoiceInfo>> GetListHoaDonCanTai(string mstcongty, string savedPath, int type)
        {
            return await Task.Run(() =>
            {
                var result = new List<InvoiceInfo>();

                try
                {
                    string folderName = type == 1 ? "HDVao" : "HDRa";
                    string currentYear = $"HD{DateTime.Now.Year}";
                    string directoryPath = Path.Combine(savedPath, currentYear, folderName, DateTime.Now.Month.ToString());

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

                    int fileIndex = 1; // Đánh dấu loại file Excel (1, 2, 3)
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

                                    if (!DateTime.TryParse(GetNLap, out DateTime getdate))
                                        continue;

                                    bool daTonTai = lookupHoaDonCT.Contains((mstnb, getSohd, getSHHD, getdate.Date, type));
                                    bool daTonTaiImport = lookupTbImport.Contains((mstnb, getSohd, getdate.Date, type));

                                    if (daTonTai || daTonTaiImport)
                                        continue;

                                    string filename = $"{getdate:yyyyMMdd}_{mstnb}_{getSohd}_{getSHHD}.xml";
                                    if (File.Exists(Path.Combine(directoryPath, filename)))
                                        continue;

                                    // ========== XÁC ĐỊNH URL THEO TYPE ==========
                                    string url;

                                    if (type == 1) // Đầu vào (HDVao)
                                    {
                                        if (fileIndex == 1 || fileIndex == 2)
                                        {
                                            // Hóa đơn điện tử có mã và không mã
                                            url = $"https://hoadondientu.gdt.gov.vn/api/query/invoices/export-xml?nbmst={mstnb}&khhdon={getSHHD}&shdon={getSohd}&khmshdon={khhd}";
                                        }
                                        else // fileIndex == 3
                                        {
                                            // Hóa đơn máy tính tiền
                                            url = $"https://hoadondientu.gdt.gov.vn/api/sco-query/invoices/export-xml?nbmst={mstnb}&khhdon={getSHHD}&shdon={getSohd}&khmshdon={khhd}";
                                        }
                                    }
                                    else // Đầu ra (HDRa)
                                    {
                                        if (fileIndex == 1)
                                        {
                                            // Hóa đơn điện tử đầu ra
                                            url = $"https://hoadondientu.gdt.gov.vn/api/sco-query/invoices/export-xml?nbmst={mstnb}&khhdon={getSHHD}&shdon={getSohd}&khmshdon={khhd}";
                                        }
                                        else // fileIndex == 2
                                        {
                                            // Hóa đơn máy tính tiền đầu ra
                                            url = $"https://hoadondientu.gdt.gov.vn/api/query/invoices/export-xml?nbmst={mstnb}&khhdon={getSHHD}&shdon={getSohd}&khmshdon={khhd}";
                                        }
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
                                    {
                                        Log($"📋 Đã đọc {result.Count} hóa đơn cần tải...");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Log($"Lỗi dòng: {ex.Message}");
                                }
                            }
                        }
                        fileIndex++;
                    }
                }
                catch (Exception ex)
                {
                    Log($"❌ Lỗi GetListHoaDonCanTai: {ex.Message}");
                }

                return result;
            });
        }
        #endregion

        #region Download Invoices
        private async Task TaiHangLoatHoaDon(CompanyHttpClient client, List<InvoiceInfo> invoices, string typeName, string companyName, DataRow companyRow = null)
        {
            try
            {
                if (invoices == null || invoices.Count == 0)
                {
                    Log($"📭 {companyName}: Không có hóa đơn {typeName} để tải");
                    return;
                }

                Log($"📥 {companyName}: Bắt đầu tải {invoices.Count} hóa đơn {typeName} (song song 3 luồng)...");

                int total = invoices.Count;
                int downloaded = 0;
                int failed = 0;
                int timeoutCount = 0;
                object lockObj = new object();

                var stopwatch = Stopwatch.StartNew();

                int maxParallel = 1;
                int timeoutSeconds = int.Parse(txttimeout.Text); // Mỗi hóa đơn tối đa 15 giây
                SemaphoreSlim semaphore = new SemaphoreSlim(maxParallel);
                List<Task> tasks = new List<Task>();

                foreach (var invoice in invoices)
                {
                    await semaphore.WaitAsync();

                    tasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            // ✅ Tạo task download với timeout
                            var downloadTask = DownloadSingleInvoiceAsync(client, invoice, companyName);
                            var timeoutTask = Task.Delay(TimeSpan.FromSeconds(timeoutSeconds));

                            // ✅ Chờ task nào hoàn thành trước
                            var completedTask = await Task.WhenAny(downloadTask, timeoutTask);

                            if (completedTask == timeoutTask)
                            {
                                // ❌ Timeout - bỏ qua hóa đơn này
                                lock (lockObj)
                                {
                                    timeoutCount++;
                                    failed++;
                                    Log($"⏰ {companyName}: Timeout HĐ {invoice.Sohd} sau {timeoutSeconds}s");
                                }
                            }
                            else
                            {
                                // ✅ Download hoàn thành (hoặc thất bại)
                                bool success = await downloadTask;
                                lock (lockObj)
                                {
                                    if (success)
                                        downloaded++;
                                    else
                                        failed++;
                                }
                            }

                            // Log tiến độ
                            lock (lockObj)
                            {
                                // ✅ Log mỗi khi có 1 hóa đơn hoàn thành (bỏ qua điều kiện % 10)
                                Log($"⏳ {companyName}: Đã xử lý {downloaded + failed}/{total} hóa đơn {typeName} (✅ {downloaded} thành công, ❌ {failed} thất bại, ⏰ {timeoutCount} timeout)");

                                // ✅ Cập nhật status vào Grid mỗi khi có 1 hóa đơn
                                if (companyRow != null)
                                {
                                    string status = $"📥 {typeName}: {downloaded + failed}/{total} (✅{downloaded} ❌{failed})";
                                    UpdateStatusOnUI(companyRow, status);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            lock (lockObj)
                            {
                                failed++;
                                Log($"❌ {companyName}: Lỗi HĐ {invoice.Sohd}: {ex.Message}");
                            }
                        }
                        finally
                        {
                            semaphore.Release();
                        }
                    }));
                }

                // ============================================================
                // ✅ CÁCH 1: KHÔNG DÙNG WhenAll - Chờ từng task với timeout
                // ============================================================
                Log($"⏳ {companyName}: Đang chờ các task hoàn thành (tối đa {timeoutSeconds + 5}s mỗi task)...");

                foreach (var task in tasks)
                {
                    try
                    {
                        // ✅ Chờ mỗi task tối đa timeoutSeconds + 5 giây
                        await Task.WhenAny(task, Task.Delay(TimeSpan.FromSeconds(timeoutSeconds + 5)));

                        // Kiểm tra nếu task chưa hoàn thành thì bỏ qua
                        if (!task.IsCompleted)
                        {
                            Log($"⏰ {companyName}: Bỏ qua task chưa hoàn thành (timeout tổng)");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Bỏ qua lỗi
                        Log($"⚠️ {companyName}: Lỗi khi chờ task: {ex.Message}");
                    }
                }

                stopwatch.Stop();
                string finalStatus = $"✅ Đã tải {typeName}: {downloaded}/{total} (✅{downloaded} ❌{failed} ⏰{timeoutCount}) - {stopwatch.Elapsed.TotalSeconds:F1}s";
                Log($"✅ {companyName}: Hoàn thành tải {typeName}! Đã tải: {downloaded}/{total}, thất bại: {failed}, timeout: {timeoutCount}, thời gian: {stopwatch.Elapsed.TotalSeconds:F1}s");

                if (companyRow != null)
                {
                    UpdateStatusOnUI(companyRow, finalStatus);
                }
            }
            catch (Exception ex)
            {
                Log($"❌ {companyName}: Lỗi TaiHangLoatHoaDon {typeName}: {ex.Message}");
            }
        }
        private async Task<bool> DownloadSingleInvoiceAsync(CompanyHttpClient client, InvoiceInfo invoice, string companyName)
        {
            if (invoice == null) return false;

            int maxRetry = 1;
            int retryCount = 0;

            while (retryCount < maxRetry)
            {
                retryCount++;

                try
                {
                    string filename = $"{invoice.NLap:yyyyMMdd}_{invoice.Mst}_{invoice.Sohd}_{invoice.SHHD}.zip";
                    string path = Path.Combine(invoice.DirectoryPath, filename);
                    string pathxml = Path.Combine(invoice.DirectoryPath, filename.Replace(".zip", ".xml"));

                    if (File.Exists(path) || File.Exists(pathxml))
                    {
                        return true;
                    }

                    bool isDownloaded = await client.DownloadFileAsync(invoice.Url, path,int.Parse(txttimeout.Text));

                    if (isDownloaded)
                    {
                        ExtractZipXML(path);
                        Log($"✅ {companyName}: Tải HĐ {invoice.Sohd} thành công");
                        return true;
                    }
                    else
                    {
                        if (retryCount < maxRetry)
                        {
                            Log($"🔄 {companyName}: Thử lại HĐ {invoice.Sohd} lần {retryCount + 1}/{maxRetry}...");
                            await Task.Delay(500);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (retryCount < maxRetry)
                    {
                        Log($"⚠️ {companyName}: Lỗi HĐ {invoice.Sohd} lần {retryCount}: {ex.Message}");
                        await Task.Delay(500);
                    }
                    else
                    {
                        Log($"❌ {companyName}: HĐ {invoice.Sohd} thất bại sau {maxRetry} lần: {ex.Message}");

                    }
                }
            }

            // Nếu thất bại, thử lấy KNM XML
            if (invoice.Type == 1)
            {
                await GetKNMXMLAsync(invoice.Mst, invoice.SHHD, invoice.Sohd, client, invoice.NLap, invoice.DirectoryPath, invoice.Sohd, companyName);
            }

            return false;
        }

        public async Task GetKNMXMLAsync(string nbmst, string khhdon, string shdon, CompanyHttpClient client, DateTime GetNLap, string path, string filename, string companyName)
        {
            string url = $"https://hoadondientu.gdt.gov.vn/api/query/invoices/detail?nbmst={nbmst}&khhdon={khhdon}&shdon={shdon}&khmshdon=1";

            try
            {
                string responseBody = await client.GetStringAsync(url);
                var rootObject = JsonConvert.DeserializeObject<Invoice>(responseBody);
                TaoFileXmlChiCoDLHDon(path, filename, rootObject, GetNLap);
                Log($"✅ {companyName}: Đã tạo KNM XML cho HĐ {shdon}");
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
            if (string.IsNullOrEmpty(invoiceNumber))
                return invoiceNumber;
            return Regex.Replace(invoiceNumber, "^0+", "");
        }

        private static void ExtractZipXML(string path)
        {
            try
            {
                Application.DoEvents();
                string rootPath = Path.GetDirectoryName(path);
                string getnamefile = Path.GetFileNameWithoutExtension(path);
                string directoryPath = rootPath + @"\Giainen" + "_" + getnamefile;

                ZipFile.ExtractToDirectory(path, directoryPath);

                var files = Directory.GetFiles(directoryPath, "invoice.html", SearchOption.AllDirectories);
                string targetFilePath = Path.Combine(rootPath, getnamefile + ".html");
                if (files.Any())
                    File.Move(files.FirstOrDefault(), targetFilePath);

                var filesxml = Directory.GetFiles(directoryPath, "invoice.xml", SearchOption.AllDirectories);
                string targetFilePathxml = Path.Combine(rootPath, getnamefile + ".xml");
                if (filesxml.Any())
                    File.Move(filesxml.FirstOrDefault(), targetFilePathxml);

                File.Delete(path);
                Directory.Delete(directoryPath, true);
            }
            catch (Exception ex)
            {
                Log($"Lỗi giải nén: {ex.Message}");

            }
        }
        #endregion

        #region XML File Creation
        public static string TaoFileXmlChiCoDLHDon(string folderPath, string fileName, Invoice Invoice, DateTime NLap)
        {
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            fileName = $"{fileName}_KNM.xml";
            string fullPath = Path.Combine(folderPath, fileName);

            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("HDon");
            doc.AppendChild(root);

            XmlElement dlHDon = doc.CreateElement("DLHDon");
            dlHDon.SetAttribute("Id", "DuLieuKy");
            root.AppendChild(dlHDon);

            // TTChung
            XmlElement ttChung = doc.CreateElement("TTChung");
            dlHDon.AppendChild(ttChung);

            ThemPhanTu(doc, ttChung, "PBan", $"{Invoice.Pban}");
            ThemPhanTu(doc, ttChung, "THDon", $"Hóa đơn GTGT");
            ThemPhanTu(doc, ttChung, "KHMSHDon", $"{Invoice.Khmshdon}");
            ThemPhanTu(doc, ttChung, "KHHDon", $"{Invoice.Khhdon}");
            ThemPhanTu(doc, ttChung, "SHDon", $"{Invoice.Shdon}");
            ThemPhanTu(doc, ttChung, "NLap", $"{NLap}");
            ThemPhanTu(doc, ttChung, "HDCTTChinh", "0");
            ThemPhanTu(doc, ttChung, "DVTTe", $"{Invoice.Dvtte}");
            ThemPhanTu(doc, ttChung, "TGia", $"{Invoice.Tgia}");
            ThemPhanTu(doc, ttChung, "HTTToan", $"{Invoice.Thtttoan}");
            ThemPhanTu(doc, ttChung, "MSTTCGP", $"{Invoice.Msttcgp}");

            // TTKhac trong TTChung
            XmlElement ttKhacChung = doc.CreateElement("TTKhac");
            ttKhacChung.AppendChild(TaoTTin(doc, "Extra", "string", ""));
            ttChung.AppendChild(ttKhacChung);

            // NDHDon
            XmlElement ndHDon = doc.CreateElement("NDHDon");
            dlHDon.AppendChild(ndHDon);

            // NBan
            XmlElement nBan = doc.CreateElement("NBan");
            ThemPhanTu(doc, nBan, "Ten", $"{Invoice.Nbten}");
            ThemPhanTu(doc, nBan, "MST", $"{Invoice.Nbmst}");
            ThemPhanTu(doc, nBan, "DChi", $"{Invoice.Nbdchi}");
            ThemPhanTu(doc, nBan, "SDThoai", $"{Invoice.Nbsdthoai}");
            ndHDon.AppendChild(nBan);

            // NMua
            XmlElement nMua = doc.CreateElement("NMua");
            ThemPhanTu(doc, nMua, "Ten", $"{Invoice.Nmten}");
            ThemPhanTu(doc, nMua, "MST", $"{Invoice.Nmmst}");
            ThemPhanTu(doc, nMua, "DChi", $"{Invoice.Nmdchi}");
            ThemPhanTu(doc, nMua, "MKHang", $"{Invoice.Mkhang}");
            ThemPhanTu(doc, nMua, "HVTNMHang", "");
            ndHDon.AppendChild(nMua);

            // DSHHDVu
            XmlElement dsHHDVu = doc.CreateElement("DSHHDVu");
            ndHDon.AppendChild(dsHHDVu);

            int stt = 1;
            if (Invoice.Hdhhdvu != null)
            {
                foreach (var dt in Invoice.Hdhhdvu.ToList())
                {
                    TaoHangHoa(doc, dsHHDVu, "0", $"{stt}", !string.IsNullOrEmpty(dt.Ten) ? $"{dt.Ten}" : "Hoá đơn không nhận mã",
                        $"{dt.Dvtinh}", $"{dt.Sluong}", $"{dt.Dgia}", $"{dt.Tsuat.Value * 100}", $"{dt.Thtien}",
                        new[] { ("Amount", "numeric", $"{dt.Thtien}"), ("VATAmount", "numeric", "0") });
                    stt++;
                }
            }

            // TToan
            XmlElement tToan = doc.CreateElement("TToan");
            ndHDon.AppendChild(tToan);

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
                Encoding = System.Text.Encoding.UTF8
            };

            using (XmlWriter writer = XmlWriter.Create(fullPath, settings))
            {
                doc.Save(writer);
            }

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

        static void TaoHangHoa(XmlDocument doc, XmlElement ds, string tChat, string stt, string tenHang, string dvTinh,
            string sl, string dGia, string tSuat, string thTien, (string truong, string kieu, string giaTri)[] extra)
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

        #region Grid Setup
        private void SetupGridCheckBox()
        {
            RepositoryItemCheckEdit checkEdit = new RepositoryItemCheckEdit();
            checkEdit.ValueChecked = 1;
            checkEdit.ValueUnchecked = 0;
            gridView1.Columns["IsRun"].ColumnEdit = checkEdit;
        }

        private void gridView1_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            try
            {
                var view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
                if (view == null) return;

                // Kiểm tra hợp lệ
                if (e.RowHandle < 0 || e.Column == null) return;

                // Lấy RowHandle và Column
                int rowHandle = e.RowHandle;
                string columnName = e.Column.FieldName;
                object newValue = e.Value;

                // Lấy dòng dữ liệu dưới dạng DataRowView
                var rowData = view.GetRow(rowHandle) as DataRowView;
                if (rowData == null) return;

                // Lấy ID từ dòng hiện tại
                int id = Convert.ToInt32(rowData["ID"] ?? 0);
                string name = rowData["Name"]?.ToString() ?? "Unknown";

                // Log
                Log($"🔄 [{name}] {columnName} = {newValue}");

                // Cập nhật database
                string query = $"UPDATE tbCompany SET {columnName} = ? WHERE ID = ?";
                var parameters = new OleDbParameter[]
                {
            new OleDbParameter("?", newValue),
            new OleDbParameter("?", id)
                };

                int rowsAffected = ExecuteQueryResult(query, parameters);

                string qr = @"SELECT * FROM tbCompany where  Saoviet = ?  order by STT";
                string computerName = Environment.MachineName;
                tbCompany = ExecuteQuery(query, new OleDbParameter("?", computerName));
                gridControl1.DataSource = tbCompany;
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi CellValueChanged: {ex.Message}");
            }
        }
        #endregion

        #region Logging
        // ✅ Log static
        private static Action<string> _logAction;

        private static void Log(string message)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} - {message}");

            // Gọi action nếu đã được khởi tạo
            if (_logAction != null)
            {
                _logAction(message);
            }
        }
        private void UpdateRichTextBox(string message)
        {
            try
            {
                // Kiểm tra richTextBox1 có tồn tại và chưa bị dispose
                if (richTextBox1 == null || richTextBox1.IsDisposed)
                    return;

                if (richTextBox1.InvokeRequired)
                {
                    // ✅ Gọi Invoke để chạy trên UI thread
                    richTextBox1.Invoke(new Action(() =>
                    {
                        if (!richTextBox1.IsDisposed)
                        {
                            richTextBox1.Text += $"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}";
                            richTextBox1.ScrollToCaret();
                        }
                    }));
                }
                else
                {
                    // Đã ở UI thread
                    richTextBox1.Text += $"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}";
                    richTextBox1.ScrollToCaret();
                }
            }
            catch (ObjectDisposedException)
            {
                // Form đã đóng, bỏ qua
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi UpdateRichTextBox: {ex.Message}");
            }
        }
        private void LogToRichTextBox(string message)
        {
            if (richTextBox1.InvokeRequired)
            {
                richTextBox1.Invoke(new Action(() => LogToRichTextBox(message)));
                return;
            }
            richTextBox1.ScrollToCaret();
        }
        #endregion

        #region Invoice Class (for JSON deserialization)
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
            public decimal Tgtcthue { get; set; }
            public decimal Tgtthue { get; set; }
            public decimal Ttcktmai { get; set; }
            public decimal Tgtttbso { get; set; }
            public string Tgtttbchu { get; set; }
            public decimal? Tgtphi { get; set; }
        }

        public class HangHoa
        {
            public string Ten { get; set; }
            public string Dvtinh { get; set; }
            public decimal Sluong { get; set; }
            public decimal Dgia { get; set; }
            public decimal? Tsuat { get; set; }
            public decimal Thtien { get; set; }
        }
        #endregion

        private void labelControl1_Click(object sender, EventArgs e)
        {

        }

        private void textEdit4_EditValueChanged(object sender, EventArgs e)
        {
            string computerName = Environment.MachineName;

            string query = "UPDATE tbsetting SET Soluongtai = ? where Saoviet=?";
            // Khai báo mảng tham số với đủ 10 tham số
            OleDbParameter[] parameters = new OleDbParameter[]
            {
                  new OleDbParameter("?", txtSoluongtai.Text),
                  new OleDbParameter("?", computerName)
            };

            // Thực thi truy vấn và lấy kết quả
            int a = ExecuteQueryResult(query, parameters);
        }

        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit1.Checked)
            {
                AddToStartup();
            }
            else
            {
                RemoveFromStartup();
            }
        }
        public void AddToStartup()
        {
            try
            {
                // 1. Lấy đường dẫn file EXE
                string exePath = Application.ExecutablePath;

                // 2. Kiểm tra file tồn tại
                if (!File.Exists(exePath))
                {
                    XtraMessageBox.Show($"❌ File EXE không tồn tại:\n{exePath}");
                    return;
                }

                // 3. Lấy tên hiển thị
                string appName = "Saoviet Auto";

                // 4. Mở Registry Run
                using (RegistryKey rk = Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (rk == null)
                    {
                        XtraMessageBox.Show("❌ Không thể mở Registry!");
                        return;
                    }

                    // ✅ Kiểm tra đã đăng ký chưa
                    string existingValue = rk.GetValue(appName) as string;

                    // Bọc trong dấu ngoặc kép + thêm tham số -autostart
                    string exeWithArgs = $"\"{exePath}\" -autostart";

                    // ✅ Nếu đã đăng ký và đúng đường dẫn thì không làm gì
                    if (!string.IsNullOrEmpty(existingValue) && existingValue == exeWithArgs)
                    {
                        //XtraMessageBox.Show($"✅ Ứng dụng đã được đăng ký Startup!\n\n" +
                        //                   $"Tên: {appName}\n" +
                        //                   $"Đường dẫn: {existingValue}",
                        //                   "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 6. Xóa entry cũ nếu có
                    if (rk.GetValue(appName) != null)
                    {
                        rk.DeleteValue(appName);
                    }

                    // 7. Thêm entry mới
                    rk.SetValue(appName, exeWithArgs);

                    // ✅ Cập nhật database
                    string computerName = Environment.MachineName;
                    int autoTaiValue = 1;
                    try
                    {
                        // Kiểm tra tồn tại
                        string checkQuery = @"SELECT COUNT(*) FROM tbsetting WHERE Saoviet = ?";
                        DataTable dt = ExecuteQuery(checkQuery, new OleDbParameter("?", computerName));

                        int count = Convert.ToInt32(dt.Rows[0][0]);
                        int rowsAffected = 0;

                        if (count > 0)
                        {
                            // UPDATE
                            string query = @"UPDATE tbsetting SET AutoTai = ? WHERE Saoviet = ?";
                            var parameters = new OleDbParameter[]
                            {
                new OleDbParameter("?", autoTaiValue),
                new OleDbParameter("?", computerName)
                            };
                            rowsAffected = ExecuteQueryResult(query, parameters);
                            Log($"✅ UPDATE AutoTai = {autoTaiValue} cho Saoviet = {computerName}");
                        }
                        else
                        {
                            // INSERT
                            string query = @"INSERT INTO tbsetting (Saoviet, AutoTai,Soluongtai,Timeout) VALUES (?, ?,?,?)";
                            var parameters = new OleDbParameter[]
                            {
                                new OleDbParameter("?", computerName),
                                new OleDbParameter("?", autoTaiValue),
                                new OleDbParameter("?", txtSoluongtai.Text),
                                 new OleDbParameter("?", txttimeout.Text)
                            };
                            rowsAffected = ExecuteQueryResult(query, parameters);
                            Log($"✅ INSERT bản ghi mới cho Saoviet = {computerName}, AutoTai = {autoTaiValue}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"❌ Lỗi SetAutoTaiForComputer: {ex.Message}");
                    } 

                    XtraMessageBox.Show($"✅ Đã đăng ký Startup!\n\n" +
                                       $"Tên: {appName}\n" +
                                       $"Đường dẫn: {exeWithArgs}\n\n" +
                                       $"📌 Khi khởi động Windows, app sẽ tự chạy và tự động xử lý.",
                                       "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show($"❌ Lỗi: {ex.Message}");
            }
        }
        public void RemoveFromStartup()
        {
            try
            {
                string appName = "Saoviet Auto";

                RegistryKey rk = Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);

                // Kiểm tra và xoá
                if (rk.GetValue(appName) != null)
                {
                    rk.DeleteValue(appName, false);
                    
                }
            }
            catch (Exception ex)
            {
              
            }
        }

        private void gridControl1_Click(object sender, EventArgs e)
        {

        }

        private void chkMoc1_CheckedChanged(object sender, EventArgs e)
        {
            txtBlock1.Enabled=chkMoc1.Checked;
            if (chkMoc1.Checked == false)
            {
                txtBlock1.Text = "";
            }
        }

        private void chkMoc2_CheckedChanged(object sender, EventArgs e)
        {
            txtBlock2.Enabled = chkMoc2.Checked;
            if (chkMoc2.Checked == false)
            {
                txtBlock2.Text = "";
            }
        }

        private void chkMoc3_CheckedChanged(object sender, EventArgs e)
        {
            txtBlock3.Enabled = chkMoc3.Checked;
            if (chkMoc3.Checked == false)
            {
                txtBlock3.Text = "";
            }
        }

        private void txtBlock1_EditValueChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBlock1.Text))
            {
                string computerName = Environment.MachineName;

                string query = "UPDATE tbsetting SET Block1 = ? where Saoviet=?";
                // Khai báo mảng tham số với đủ 10 tham số
                OleDbParameter[] parameters = new OleDbParameter[]
                {
        new OleDbParameter("?", txtBlock1.Text),
          new OleDbParameter("?",computerName)
                };

                // Thực thi truy vấn và lấy kết quả
                int a = ExecuteQueryResult(query, parameters);
                ScheduleHelper ScheduleHelper = new ScheduleHelper();
                DeleteSchedule($"Block_L1");
            }
        }

        private void txtBlock2_EditValueChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBlock2.Text))
            {
                string computerName = Environment.MachineName;
                string query = "UPDATE tbsetting SET Block2 = ? where Saoviet=?";
                // Khai báo mảng tham số với đủ 10 tham số
                OleDbParameter[] parameters = new OleDbParameter[]
                {
                  new OleDbParameter("?", txtBlock2.Text),
                  new OleDbParameter("?",computerName)
                };

                // Thực thi truy vấn và lấy kết quả
                int a = ExecuteQueryResult(query, parameters);
                ScheduleHelper ScheduleHelper = new ScheduleHelper();
                DeleteSchedule($"Block_L2");
            }
        }

        private void txtBlock3_EditValueChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBlock1.Text))
            {
                string computerName = Environment.MachineName;
                string query = "UPDATE tbsetting SET Block3 = ? where Saoviet=?";
                // Khai báo mảng tham số với đủ 10 tham số
                OleDbParameter[] parameters = new OleDbParameter[]
                {
                  new OleDbParameter("?", txtBlock3.Text),
                  new OleDbParameter("?",computerName)
                };

                // Thực thi truy vấn và lấy kết quả
                int a = ExecuteQueryResult(query, parameters);
                ScheduleHelper ScheduleHelper = new ScheduleHelper();
                DeleteSchedule($"Block_L3");
            }
        }
        public static void DeleteSchedule(string taskName)
        {
            try
            {
                string cmd = $"schtasks /delete /tn \"{taskName}\" /f";
                Process.Start("cmd", "/c " + cmd).WaitForExit();
                XtraMessageBox.Show($"✅ Đã xóa lịch '{taskName}'!");
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show($"❌ Lỗi xóa: {ex.Message}");
            }
        }

        private void txtBlock1_Validated(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra tồn tại
                string computerName = Environment.MachineName;
                string checkQuery = @"SELECT COUNT(*) FROM tbsetting WHERE Saoviet = ?";
                DataTable dt = ExecuteQuery(checkQuery, new OleDbParameter("?", computerName));

                int count = Convert.ToInt32(dt.Rows[0][0]);
                int rowsAffected = 0;

                if (count > 0)
                {
                    // UPDATE
                    string query = @"UPDATE tbsetting SET Block1 = ? WHERE Saoviet = ?";
                    var parameters = new OleDbParameter[]
                    {
                new OleDbParameter("?", txtBlock1.Text),
                new OleDbParameter("?", computerName)
                    };
                    rowsAffected = ExecuteQueryResult(query, parameters); 
                }
                else
                {
                    // INSERT 
                    string query = @"INSERT INTO tbsetting (Saoviet, Block1,Soluongtai,Timeout) VALUES (?, ?,?,?)";
                    var parameters = new OleDbParameter[]
                    {
                                new OleDbParameter("?", computerName),
                                  new OleDbParameter("?", txtBlock1.Text),
                                new OleDbParameter("?", txtSoluongtai.Text),
                                 new OleDbParameter("?", txttimeout.Text)
                    };
                    rowsAffected = ExecuteQueryResult(query, parameters);
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi SetAutoTaiForComputer: {ex.Message}");
            }
            ScheduleHelper ScheduleHelper = new ScheduleHelper();
            DateTime time = DateTime.Parse(txtBlock1.Text);

            int hour = time.Hour;   // 10
            int minute = time.Minute; // 30
            try
            {
                CreateSchedule($"Block_L1", hour, minute);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        public static void CreateSchedule(string taskName, int hour, int minute = 0)
        {
            try
            {
                string exePath = Application.ExecutablePath;
                string timeStr = $"{hour:D2}:{minute:D2}";

                // Xóa task cũ (nếu có)
                RunSchTasks($"/delete /tn \"{taskName}\" /f", false);

                // Chuỗi thực thi
                string tr = $"\\\"{exePath}\\\" -autostart";

                // Tạo task
                string args =
                    $"/create /tn \"{taskName}\" " +
                    $"/tr \"{tr}\" " +
                    $"/sc daily " +
                    $"/st {timeStr} " +
                    $"/f";

                // Debug xem lệnh tạo ra
                XtraMessageBox.Show(args);

                RunSchTasks(args, true);

                XtraMessageBox.Show($"Đã tạo lịch '{taskName}' lúc {timeStr} mỗi ngày.");
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
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
                {
                    throw new Exception(
                        $"ExitCode: {p.ExitCode}\r\n\r\n" +
                        $"Output:\r\n{output}\r\n\r\n" +
                        $"Error:\r\n{error}"
                    );
                }
            }
        }

        private void txttimeout_EditValueChanged(object sender, EventArgs e)
        {
            string query = "UPDATE tbsetting SET Timeout = ? where Saoviet=?";
            string computerName = Environment.MachineName;

            // Khai báo mảng tham số với đủ 10 tham số
            OleDbParameter[] parameters = new OleDbParameter[]
            {
                new OleDbParameter("?", txttimeout.Text),
                new OleDbParameter("?",computerName)
            };

            // Thực thi truy vấn và lấy kết quả
            int a = ExecuteQueryResult(query, parameters);
        }

        private void txtBlock2_Validated(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra tồn tại
                string computerName = Environment.MachineName;
                string checkQuery = @"SELECT COUNT(*) FROM tbsetting WHERE Saoviet = ?";
                DataTable dt = ExecuteQuery(checkQuery, new OleDbParameter("?", computerName));

                int count = Convert.ToInt32(dt.Rows[0][0]);
                int rowsAffected = 0;

                if (count > 0)
                {
                    // UPDATE
                    string query = @"UPDATE tbsetting SET Block2 = ? WHERE Saoviet = ?";
                    var parameters = new OleDbParameter[]
                    {
                new OleDbParameter("?", txtBlock2.Text),
                new OleDbParameter("?", computerName)
                    };
                    rowsAffected = ExecuteQueryResult(query, parameters);
                }
                else
                {
                    // INSERT 
                    string query = @"INSERT INTO tbsetting (Saoviet, Block1,Soluongtai,Timeout) VALUES (?, ?,?,?)";
                    var parameters = new OleDbParameter[]
                    {
                                new OleDbParameter("?", computerName),
                                  new OleDbParameter("?", txtBlock2.Text),
                                new OleDbParameter("?", txtSoluongtai.Text),
                                 new OleDbParameter("?", txttimeout.Text)
                    };
                    rowsAffected = ExecuteQueryResult(query, parameters);
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi SetAutoTaiForComputer: {ex.Message}");
            }
            ScheduleHelper ScheduleHelper = new ScheduleHelper();
            DateTime time = DateTime.Parse(txtBlock2.Text);

            int hour = time.Hour;   // 10
            int minute = time.Minute; // 30
            try
            {
                CreateSchedule($"Block_L2", hour, minute);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void txtBlock3_Validated(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra tồn tại
                string computerName = Environment.MachineName;
                string checkQuery = @"SELECT COUNT(*) FROM tbsetting WHERE Saoviet = ?";
                DataTable dt = ExecuteQuery(checkQuery, new OleDbParameter("?", computerName));

                int count = Convert.ToInt32(dt.Rows[0][0]);
                int rowsAffected = 0;

                if (count > 0)
                {
                    // UPDATE
                    string query = @"UPDATE tbsetting SET Block3 = ? WHERE Saoviet = ?";
                    var parameters = new OleDbParameter[]
                    {
                new OleDbParameter("?", txtBlock3.Text),
                new OleDbParameter("?", computerName)
                    };
                    rowsAffected = ExecuteQueryResult(query, parameters);
                }
                else
                {
                    // INSERT 
                    string query = @"INSERT INTO tbsetting (Saoviet, Block1,Soluongtai,Timeout) VALUES (?, ?,?,?)";
                    var parameters = new OleDbParameter[]
                    {
                                new OleDbParameter("?", computerName),
                                  new OleDbParameter("?", txtBlock3.Text),
                                new OleDbParameter("?", txtSoluongtai.Text),
                                 new OleDbParameter("?", txttimeout.Text)
                    };
                    rowsAffected = ExecuteQueryResult(query, parameters);
                }
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi SetAutoTaiForComputer: {ex.Message}");
            }
            ScheduleHelper ScheduleHelper = new ScheduleHelper();
            DateTime time = DateTime.Parse(txtBlock3.Text);

            int hour = time.Hour;   // 10
            int minute = time.Minute; // 30
            try
            {
                CreateSchedule($"Block_L3", hour, minute);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }


        #region Dochoadon
        public class VatTuInfo
        {
            public string Ma { get; set; }
            public string DonViTinh { get; set; }
        }
        private Dictionary<string, string> _lookupByTenChinh;
        private Dictionary<string, VatTuInfo> _lookupByTenChinhs;
        public Dictionary<string, string> _lookupByTenPhu;
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
        public List<VatTu> lstvt = new List<VatTu>();
        public DataTable existingTbChungtu;
        DataTable existingTbHeThongTK;
        public DataTable existingTbHoadon;
        DataTable tbNhapkhonguyenlieu;
        DataTable tbTonkho;
        DataTable tbNhapkhotp;
        DataTable tbRegister;
        DataTable tbLicense;
        DataTable ListPhanloaiVattu;
        private Dictionary<string, (string TenVattu, string TenVattu2, string DonVi, double Dongia, double SoLuong)> vatTuLookup;

        // ==================== TOÀN CỤC (chỉ khai báo 1 lần) ====================
        private readonly Dictionary<string, (string SoHieu, double Percent)> _cacheToanCuc
            = new Dictionary<string, (string SoHieu, double Percent)>(StringComparer.OrdinalIgnoreCase);
        public async Task<List<VatTu>> LoadDataVattuAsync(string conectionst)
        {
            // Hiển thị popup loading
            List<VatTu> lstVattu = new List<VatTu>();

            try
            {
                // 1. Lấy danh sách VatTu từ database

                var queryVatTu = @"SELECT * FROM Vattu";
                var ListVattu = await Task.Run(() => ExecuteQuery2(queryVatTu, conectionst, null));
                var queryMaphanloai = @"SELECT * FROM PhanLoaiVattu";
                ListPhanloaiVattu = await Task.Run(() => ExecuteQuery2(queryMaphanloai, conectionst, null));

                // 2. Chuyển đổi chuỗi VNI sang Unicode (nếu cần)
                foreach (DataRow item in ListVattu.Rows)
                {
                    item["TenVattu"] = Helpers.ConvertVniToUnicode(item["TenVattu"].ToString());
                    item["TenVattu2"] = Helpers.ConvertVniToUnicode(item["TenVattu2"].ToString());
                    item["DonVi"] = Helpers.ConvertVniToUnicode(item["DonVi"].ToString());
                }

                // 3. Gom nhóm tất cả MaVatTu để query TonKho 1 lần duy nhất (Batch Query)
                var maVatTuList = ListVattu.Rows
                    .Cast<DataRow>()
                    .Select(row => int.Parse(row["MaSo"].ToString()))
                    .Distinct()
                    .ToList();
                if (maVatTuList.Count == 0)
                    return new List<VatTu>();
                // 4. Lấy dữ liệu TonKho theo danh sách MaVatTu đã gom nhóm
                var queryTonKhoBatch = @"SELECT * FROM TonKho WHERE MaVatTu IN (" +
                                       string.Join(",", maVatTuList) + ")";
                var allTonKho = await Task.Run(() => ExecuteQuery2(queryTonKhoBatch, conectionst, null));

                // 5. Chuyển dữ liệu TonKho thành Dictionary để truy cập nhanh bằng MaVatTu
                var tonKhoDict = allTonKho.Rows
                    .Cast<DataRow>()
                    .GroupBy(row => int.Parse(row["MaVatTu"].ToString()))
                    .ToDictionary(group => group.Key, group => group.First());

                // 6. Xử lý từng VatTu và ánh xạ dữ liệu TonKho tương ứng
                List<Task<VatTu>> vatTuTasks = new List<Task<VatTu>>();


                foreach (DataRow item in ListVattu.Rows)
                {
                    try
                    {
                        // Lưu trữ dữ liệu cần thiết để tránh closure issues
                        var maSo = int.Parse(item["MaSo"].ToString());
                        var maPhanLoai = int.Parse(item["MaPhanLoai"].ToString());
                        var tenVattu = item["TenVattu"].ToString();
                        var tenVattu2 = item["TenVattu2"].ToString();
                        var soHieu = item["SoHieu"].ToString();
                        var donVi = item["DonVi"].ToString();
                        var ghiChu = item["GhiChu"].ToString();
                        var tenMaPhanLoai = ListPhanloaiVattu.AsEnumerable()
                            .Where(m => m["MaSo"].ToString() == item["MaPhanLoai"].ToString())
                            .FirstOrDefault()?["TenPhanLoai"].ToString() ?? string.Empty;
                        var ptgb = item["PTGB"].ToString();

                        var task = Task.Run(() =>
                        {
                            var VatTu = new VatTu
                            {
                                MaSo = maSo,
                                MaPhanLoai = maPhanLoai,
                                TenVattu = tenVattu,
                                TenVattu2 = tenVattu2,
                                SoHieu = soHieu,
                                DonVi = donVi,
                                GhiChu = ghiChu,
                                TenMaPhanLoai = tenMaPhanLoai,
                                PTGB = ptgb,
                            };

                            // Kiểm tra và lấy dữ liệu từ TonKho (nếu có)
                            if (tonKhoDict.TryGetValue(VatTu.MaSo, out DataRow tonKhoRow))
                            {
                                int cnt = 12;

                                // Lấy số lượng và thành tiền
                                var soluong = tonKhoRow["Luong_" + cnt] != DBNull.Value
                                    ? double.Parse(tonKhoRow["Luong_" + cnt].ToString())
                                    : 0;
                                VatTu.SoLuong = soluong;
                                //Tìm số lượng thông qua tbchungtu


                                var thanhtien = tonKhoRow["Tien_" + cnt] != DBNull.Value
                                   ? double.Parse(tonKhoRow["Tien_" + cnt].ToString())
                                   : 0;
                                VatTu.ThanhTien = thanhtien;

                                // Tính đơn giá nếu có dữ liệu
                                if (soluong != 0 && thanhtien != 0)
                                {
                                    VatTu.Dongia = thanhtien / soluong;
                                }
                                try
                                {
                                    if (existingTbChungtu != null)
                                    {
                                        var findlstct = existingTbChungtu.AsEnumerable().Where(m => int.Parse(m["MaVattu"].ToString()) == VatTu.MaSo && double.Parse(m["SoPS"].ToString()) != 0 && double.Parse(m["SoPS2Co"].ToString()) != 0 && !m["SoHieu"].ToString().Contains("V")).ToList().LastOrDefault();
                                        if (findlstct != null)
                                        {
                                            if (findlstct["MaSo"].ToString() == "174028")
                                            {
                                                int dd = 10;
                                            }
                                            double SoPS2Co = double.Parse(findlstct["SoPS2Co"].ToString());
                                            double SoPS = double.Parse(findlstct["SoPS"].ToString());
                                            if (SoPS2Co > 0)
                                                VatTu.Dongia2 = Math.Round(findlstct.Field<double>("SoPS") / SoPS2Co);
                                            else
                                                VatTu.Dongia2 = 0;
                                        }
                                        //if (VatTu.Dongia2 != 0)
                                        //    VatTu.Dongia = VatTu.Dongia2;
                                    }

                                }
                                catch (Exception ex)
                                {
                                    XtraMessageBox.Show($"Lỗi khi tính đơn giá vật tư {VatTu.Dongia2}: {ex.Message}");
                                }
                            }
                            return VatTu;
                        });

                        vatTuTasks.Add(task);
                    }
                    catch (Exception ex)
                    {
                        XtraMessageBox.Show($"Lỗi khi khởi tạo Task vật tư: {ex.Message}");
                    }
                }

                // 7. Đợi tất cả các Task hoàn thành và thêm vào danh sách kết quả
                try
                {
                    var vatTus = await Task.WhenAll(vatTuTasks);
                    lstVattu.AddRange(vatTus.Where(v => v != null));
                }
                catch (AggregateException aggEx)
                {
                    foreach (var ex in aggEx.InnerExceptions)
                    {
                        XtraMessageBox.Show($"Lỗi khi xử lý vật tư: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show($"Lỗi khi xử lý vật tư: {ex.Message}");
                }


                //  XtraMessageBox.Show("Load vattu thanh cong");
            }
            catch (Exception ex)
            {
                // Xử lý lỗi (có thể log hoặc hiển thị thông báo)
                Console.WriteLine($"Lỗi khi tải dữ liệu: {ex.Message} ");
                throw; // Re-throw nếu cần thiết
            }
            finally
            {
                // Đóng popup loading chỉ khi mọi thứ đã hoàn tất
            }
            // BuildFastLookup();

            _lookupByTenChinh = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _lookupByTenPhu = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _lookupByTenChinhs = new Dictionary<string, VatTuInfo>(StringComparer.OrdinalIgnoreCase);

            vatTuLookup = lstVattu
           .ToDictionary(v => v.SoHieu, v => (v.TenVattu, v.TenVattu2, v.DonVi, v.Dongia, v.SoLuong));
            foreach (var kvp in vatTuLookup)
            {
                string sohieu = kvp.Key;
                if (sohieu == "TBC-001")
                {
                    int test = 10;
                }

                // DÙNG CHÍNH XÁC HÀM NormalizeForLookup
                string key1 = Helpers.NormalizeVietnameseString(kvp.Value.TenVattu.Trim());

                VietnameseProductMatcher vietnameseProductMatcher = new VietnameseProductMatcher();
                key1 = vietnameseProductMatcher.NormalizeVietnameseProduct(key1);

                if (!string.IsNullOrEmpty(key1))
                {
                    _lookupByTenChinh[key1] = sohieu;
                    _lookupByTenChinhs[key1] = new VatTuInfo
                    {
                        Ma = sohieu,
                        DonViTinh = kvp.Value.DonVi
                    };
                }

                if (!string.IsNullOrEmpty(kvp.Value.TenVattu2))
                {
                    string key2 = Helpers.NormalizeVietnameseString(kvp.Value.TenVattu2.Trim());
                    if (!string.IsNullOrEmpty(key2))
                        _lookupByTenPhu[key2] = sohieu;
                }
            }
            InitializeVatTuOptimization();
            return lstVattu;
        }
        private Dictionary<string, (string TenChuan, string TenPhuChuan, string QuyCach, string DonVi, double Dongia, double soluong)> _optimizedVatTu;

        private void InitializeVatTuOptimization()
        {
            _optimizedVatTu = new Dictionary<string, (string, string, string, string, double, double)>();
            Regex regex = new Regex(@"(\d+(g|ml|L|kg)|x\d+|(\d+\s*cái))", RegexOptions.IgnoreCase);

            foreach (var item in vatTuLookup)
            {
                string ten1 = Helpers.NormalizeVietnameseString(item.Value.TenVattu);
                string ten2 = Helpers.NormalizeVietnameseString(item.Value.TenVattu2);
                string quyCach = regex.Match(ten1).Value;

                _optimizedVatTu[item.Key] = (ten1, ten2, quyCach, item.Value.DonVi, item.Value.Dongia, item.Value.SoLuong);
            }
        }
        private async Task XulylietkeHoaDon(int type,string connectionString2)
        {
            lstvt = await LoadDataVattuAsync(connectionString2);
            string pathType = type == 1 ? "HDVao" : "HDRa";
            int fromMonth = DateTime.Now.Month;
            int toMonth = DateTime.Now.Month;
            string pathYear = $"HD{DateTime.Now.Year}";
            // 2. Gom tất cả file XML
            List<string> allFiles = new List<string>();
            for (int m = fromMonth; m <= toMonth; m++)
            {
                string query = @"SELECT * FROM tbRegister";
                var getResgistry = ExecuteQuery2(query, connectionString2);
                string savedPath = getResgistry.Rows[0]["Hoadonpath"].ToString();
                string monthFolder = Path.Combine(savedPath, pathYear, pathType, m.ToString());
                if (Directory.Exists(monthFolder))
                {
                    var filesInMonth = Directory.GetFiles(monthFolder, "*.xml", SearchOption.TopDirectoryOnly);
                    allFiles.AddRange(filesInMonth);
                }
            }
            List<TbImport> allInvoicesToSave = new List<TbImport>();
            int batchSize = 10;
            for (int i = 0; i < allFiles.Count; i += batchSize)
            {
                var batch = allFiles.Skip(i).Take(batchSize);

                // Tạo các task đọc file
                var tasks = batch.Select(file => DocfileXmlOne(file, 1, type));

                // --- ĐOẠN QUAN TRỌNG: Hứng kết quả từ các file vừa đọc ---
                TbImport[] results = await Task.WhenAll(tasks);

                foreach (var item in results)
                {
                    if (item != null) // Chỉ lấy những file bóc tách thành công và không trùng
                    {
                        allInvoicesToSave.Add(item);

                        // Cập nhật vào danh sách hiển thị trên giao diện (Grid)
                        if (type == 1) lstdsVao.Add(item);
                        else lstdsRa.Add(item);
                    }
                }

                // Cập nhật UI
                Application.DoEvents();
            }
            if (allInvoicesToSave.Count > 0)
            {
                await SaveAllInvoicesBulk(allInvoicesToSave, type == 1 ? 1 : 2, connectionString2);

            }
        }
        private List<TbImport> lstdsVao = new List<TbImport>();

        private List<TbImport> lstdsRa = new List<TbImport>();
        private async Task<TbImport> DocfileXmlOne(string pathXml, int stt, int type)
        {
            if (pathXml.Contains("427"))
            {
                int kiemtra = 10;
            }
            isAddhd = true;

            if (pathXml.Contains("html"))
                return null;

            TbImport tbImport = null;
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Parse; // Cho phép phân tích DTD
                settings.XmlResolver = null; // Ngăn chặn việc tải DTD từ bên ngoài để bảo mật và tốc độ
                settings.Async = true;
                using (var xmlReader = XmlReader.Create(pathXml, new XmlReaderSettings { Async = true }))
                {
                    var xmlDoc = new XmlDocument();
                    xmlDoc.Load(xmlReader);
                    XmlNode root = xmlDoc.DocumentElement;
                    if (root == null) return null;

                    // 1. Khai báo các Node cha để tối ưu truy vấn (tránh dùng // liên tục)
                    XmlNode ttChung = root.SelectSingleNode("//TTChung");
                    XmlNode nBan = root.SelectSingleNode("//NBan");
                    XmlNode nMua = root.SelectSingleNode("//NMua");
                    XmlNode ttToan = root.SelectSingleNode("//TToan");
                    XmlNode THDon = root.SelectSingleNode("//THDon");
                    if (ttChung == null || ttToan == null) return null;

                    tbImport = new TbImport { Path = pathXml };


                    if (Helpers.NormalizeVietnameseString(THDon.InnerText.ToLower()).Contains("hóa đơn giá trị gia tăng") || Helpers.NormalizeVietnameseString(THDon.InnerText.ToLower()).Contains("hóa đơn điện tử giá trị gia tăng"))
                    {
                        tbImport.hdon = "01";
                    }
                    if (Helpers.NormalizeVietnameseString(THDon.InnerText.ToLower()).Contains("hóa đơn bán hàng"))
                    {
                        tbImport.hdon = "02";
                    }

                    // 2. Xử lý NLap và nội dung điều chỉnh
                    if (DateTime.TryParse(ttChung.SelectSingleNode("NLap")?.InnerText, out DateTime nLap))
                        tbImport.NLap = nLap;

                    // Kiểm tra trong khoảng ngày 
                    // Nội dung điều chỉnh từ TTKhac
                    var ttKhacNodes = ttChung.SelectNodes("TTKhac/TTin");
                    if (ttKhacNodes != null)
                    {
                        foreach (XmlNode node in ttKhacNodes)
                        {
                            string dLieu = node.SelectSingleNode("DLieu")?.InnerText;
                            if (!string.IsNullOrEmpty(dLieu) && dLieu.IndexOf("điều chỉnh", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                tbImport.Noidung = Helpers.ConvertUnicodeToVni(dLieu);
                                break;
                            }
                        }
                    }

                    // 3. Thông tin số hóa đơn & Ký hiệu
                    tbImport.SHDon = Helpers.RemoveLeadingZeros(ttChung.SelectSingleNode("SHDon")?.InnerText);
                    if (tbImport.SHDon == "1" && tbImport.KHHDon == "C26MKK")
                    {
                        int test = 10;
                    }
                    tbImport.KHHDon = ttChung.SelectSingleNode("KHHDon")?.InnerText;
                    //Lấy thông tin thay thế

                    string TPhi = ttToan.SelectSingleNode("//TPhi")?.InnerText;
                    if (!string.IsNullOrEmpty(TPhi))
                    {
                        tbImport.TPhi = double.Parse(TPhi).ToString();
                    }
                    // 4. Phân loại Mua/Bán và Kiểm tra MST công ty
                    if (type == 1) // Hóa đơn đầu vào
                    {
                        GetMST();
                        // if (mstcongty != nMua?.SelectSingleNode("MST")?.InnerText && CCCD != nMua?.SelectSingleNode("MST")?.InnerText) return null;
                        tbImport.Ten = Helpers.ConvertUnicodeToVni(nBan?.SelectSingleNode("Ten")?.InnerText ?? "");
                        tbImport.Mst = nBan?.SelectSingleNode("MST")?.InnerText ?? "";
                    }
                    else // Hóa đơn đầu ra
                    {
                        //if (mstcongty != nBan?.SelectSingleNode("MST")?.InnerText)
                        //{
                        //    return null;
                        //}

                        string tenDoiTac =
  !string.IsNullOrWhiteSpace(nMua?.SelectSingleNode("Ten")?.InnerText)
      ? nMua.SelectSingleNode("Ten").InnerText
      : nMua?.SelectSingleNode("HVTNMHang")?.InnerText;

                        tbImport.Ten = !string.IsNullOrEmpty(tenDoiTac) ? Helpers.ConvertUnicodeToVni(tenDoiTac) : "";
                        tbImport.Mst = nMua?.SelectSingleNode("MST")?.InnerText ?? nMua?.SelectSingleNode("CCCDan")?.InnerText ?? "";

                        // Xử lý khách lẻ
                        if (string.IsNullOrEmpty(tbImport.Ten) || tbImport.Ten.Contains("khaùch khoâng laáy hoùa ñôn") || tbImport.Ten.Contains("Ngöôøi mua khoâng laáy hoùa ñôn") || tbImport.Ten.Contains("Khaùch leû"))
                        {
                            var kl = tbKhachhang.AsEnumerable().FirstOrDefault(m => m.Field<string>("SoHieu") == "KL");
                            if (kl != null) { tbImport.Ten = kl.Field<string>("Ten"); tbImport.Mst = "KL"; }
                        }
                    }

                    // 5. Tạo Số hiệu tự động nếu không có MST

                    //Kiểm tra tồn tại khách hàng


                    if (string.IsNullOrEmpty(tbImport.Mst) && !string.IsNullOrEmpty(tbImport.Ten))
                    {
                        var existingSoHieus = tbKhachhang.AsEnumerable().Where(m => m.Field<string>("Ten").ToLower() == tbImport.Ten.ToLower()).Select(r => r.Field<string>("SoHieu")).ToList();
                        if (existingSoHieus == null || existingSoHieus.Count == 0)
                        {
                            string sohieuBase = GenerateAbbreviation(Helpers.ConvertVniToUnicode(tbImport.Ten), existingSoHieus).ToUpper();
                            string finalSH = sohieuBase;
                            int suffix = 1;
                            while (tbKhachhang.AsEnumerable().Any(r => r.Field<string>("SoHieu") == finalSH))
                                finalSH = $"{sohieuBase}_{suffix++}";
                            tbImport.Mst = finalSH;
                        }
                        else
                        {
                            tbImport.Mst = existingSoHieus.FirstOrDefault();
                        }

                    }

                    // 6. Kiểm tra trùng hóa đơn (Cache & Database)
                    var currentList = type == 1 ? lstdsVao : lstdsRa;
                    var importList = type == 1 ? lstImportVao : lstImportRa;
                    if (currentList.Any(m => m.SHDon == tbImport.SHDon && m.NLap.Date == tbImport.NLap.Date && m.Mst == tbImport.Mst)) return null;
                    if (importList.Any(m => m.SHDon == tbImport.SHDon && m.NLap.Date == tbImport.NLap.Date)) return null;
                    if ((Kiemtrahoadon(tbImport.SHDon, tbImport.NLap, tbImport.Mst, type)) || KiemtrahoadonCT(tbImport.SHDon, tbImport.KHHDon, tbImport.NLap, tbImport.Mst, type))
                    {
                        isAddhd = false;
                        return null;
                    }

                    // 7. Khởi tạo khách hàng mới
                    if (tbImport.Mst != "KL" && !CheckExistKH(tbImport.Mst))
                    {
                        XmlNode doiTacNode = (type == 1) ? nBan : nMua;
                        string dChi = doiTacNode?.SelectSingleNode("DChi")?.InnerText ?? "";
                        string sdt = doiTacNode?.SelectSingleNode("SDThoai")?.InnerText ?? "";
                        if (!string.IsNullOrEmpty(tbImport.Mst) && !string.IsNullOrEmpty(tbImport.Ten) && !CheckExistKH(tbImport.Mst))
                            InitCustomer(type == 1 ? 2 : 3, tbImport.Mst, tbImport.Ten, dChi, tbImport.Mst, "", sdt);
                        else
                        {
                            var kl = tbKhachhang.AsEnumerable().FirstOrDefault(m => m.Field<string>("SoHieu") == "KL");
                            if (kl != null) { tbImport.Ten = kl.Field<string>("Ten"); tbImport.Mst = "KL"; }
                        }
                    }

                    // 8. Định danh tài khoản
                    string kw = type == 1 ? "Ưu tiên vào" : "Ưu tiên ra";
                    var authRow = tbDinhDanhtaikhoan.AsEnumerable().FirstOrDefault(r => r.Field<string>("KeyValue")?.Contains(kw) == true);
                    if (authRow != null)
                    {
                        tbImport.TKNo = authRow["TKNo"]?.ToString();
                        tbImport.TKCo = authRow["TKCo"]?.ToString();
                        tbImport.TkThue = authRow["TKThue"]?.ToString();
                    }
                    tbImport.Status = 0;
                    tbImport.Ngaytao = DateTime.Now.ToShortDateString();

                    // 9. Tiền thanh toán và Thuế suất
                    tbImport.TongTien = double.Parse(ttToan.SelectSingleNode("TgTTTBSo")?.InnerText ?? "0");
                    tbImport.TgTCThue = double.Parse(ttToan.SelectSingleNode("TgTCThue")?.InnerText ?? "0");
                    tbImport.TgTThue = double.Parse(ttToan.SelectSingleNode("TgTThue")?.InnerText ?? "0");

                    tbImport.Vat = 0;
                    tbImport.Vat2 = "0";
                    tbImport.Vat3 = "0";
                    var thueNodes = ttToan.SelectNodes("THTTLTSuat//LTSuat");
                    if (thueNodes != null)
                    {
                        for (int i = 0; i < thueNodes.Count; i++)
                        {
                            XmlNode n = thueNodes[i];
                            string tsStr = n.SelectSingleNode("TSuat")?.InnerText ?? "";
                            double ttien = double.Parse(n.SelectSingleNode("ThTien")?.InnerText ?? "0");
                            double tthue = Math.Round(double.Parse(n.SelectSingleNode("TThue")?.InnerText ?? "0"));
                            double vVal = (tsStr == "KCT" || tsStr == "KKKNT") ? 0 : double.Parse(tsStr.Replace("%", ""));
                            if ((tsStr == "KCT" && ttien == 0) || (tsStr == "KKKNT" && ttien == 0) || (tsStr == "0%" && ttien == 0) || ttien == 0)
                                continue;
                            if (tbImport.TgTCThue1 == 0)
                            {
                                tbImport.TgTCThue1 = ttien; tbImport.TVat = tthue; tbImport.Vat = vVal;
                            }
                            else
                            {
                                if (tbImport.TgTCThue2 == 0)
                                {
                                    tbImport.TgTCThue2 = ttien; tbImport.TVat2 = tthue; tbImport.Vat2 = vVal.ToString();
                                }
                                else
                                {
                                    if (tbImport.TgTCThue3 == 0)
                                    {
                                        tbImport.TgTCThue3 = ttien; tbImport.TVat3 = tthue; tbImport.Vat3 = vVal.ToString();
                                    }
                                }
                            }
                            //if (i == 0) { tbImport.TgTCThue1 = ttien; tbImport.TVat = tthue; tbImport.Vat = vVal; }
                            //else if (i == 1) { tbImport.TgTCThue2 = ttien; tbImport.TVat2 = tthue; tbImport.Vat2 = vVal.ToString(); }
                            //else if (i == 2) { tbImport.TgTCThue3 = ttien; tbImport.TVat3 = tthue; tbImport.Vat3 = vVal.ToString(); }
                        }
                    }
                    //Xử lý dong thuế thưa
                    //if (tbImport.Vat2 != "0" && !string.IsNullOrEmpty(tbImport.Vat2) && tbImport.Vat == 0)
                    //{
                    //    tbImport.Vat = double.Parse(tbImport.Vat2);
                    //    tbImport.TVat = tbImport.TVat2;
                    //    tbImport.Vat2 = "0";
                    //    tbImport.TVat2 = 0;
                    //}
                    // 10. Loại hóa đơn (01: GTGT, 02: Bán hàng)
                    tbImport.Khmshdon = root.SelectSingleNode("//KHMSHDon")?.InnerText;
                    string thDon = Helpers.NormalizeVietnameseString(root.SelectSingleNode("//THDon")?.InnerText?.ToLower() ?? "");
                    //      tbImport.hdon = thDon.Contains("ban hang") ? "02" : "01";

                    // 11. Chi tiết hàng hóa (HHDVu)
                    var hhdNodes = root.SelectNodes("//HHDVu");
                    cacheMatHangTrongHoaDon = new Dictionary<string, TbImportDetail>(StringComparer.OrdinalIgnoreCase);
                    double finalTotal = 0;
                    foreach (XmlNode node in hhdNodes)
                    {
                        string ten = "";
                        try
                        {
                            string tenGoc = node.SelectSingleNode("THHDVu")?.InnerText;
                            ten = tenGoc;
                            if (ten == "Hộp xích xe máy C110-SXT")
                            {
                                int test = 10;
                            }
                            if (Loaiborow(tenGoc)) continue;

                            int tchat = int.Parse(node.SelectSingleNode("TChat")?.InnerText ?? "0");
                            if (tenGoc.Contains("Chiết khấu") && tchat != 3)
                            {
                                tchat = 3;
                            }
                            bool daGiam = tenGoc.Contains("Đã giảm");
                            if (tchat == 4 && !daGiam) continue;

                            TbImportDetail dt = new TbImportDetail
                            {
                                Tchat = tchat,
                                Ten = tenGoc,
                                TKNo = tbImport.TKNo,
                                TKCo = tbImport.TKCo,
                                // Thêm kiểm tra null cho DVTinh
                                DVT = CapitalizeFirstLetters(Helpers.ConvertUnicodeToVni(node.SelectSingleNode("DVTinh")?.InnerText ?? "")),
                                // Sử dụng SafeParse để không bao giờ bị văng lỗi
                                Soluong = SafeParse(node.SelectSingleNode("SLuong")?.InnerText),
                                Dongia = SafeParse(node.SelectSingleNode("DGia")?.InnerText),
                                TTien = SafeParse(node.SelectSingleNode("ThTien")?.InnerText),
                                SoPSGoc = SafeParse(node.SelectSingleNode("ThTien")?.InnerText),
                                Vat = SafeParse(node.SelectSingleNode("TSuat")?.InnerText?.Replace("%", ""))
                            };
                            if (string.IsNullOrEmpty(dt.DVT))
                            {
                                var findvt = lstvt.FirstOrDefault(m => m.TenVattu.ToLower() == dt.Ten.ToLower());
                                if (findvt != null)
                                {
                                    dt.DVT = Helpers.ConvertUnicodeToVni(findvt.DonVi);
                                }
                            }
                            finalTotal += dt.TTien;
                            if (daGiam)
                            {
                                Match m = Regex.Match(tenGoc, @"\d{1,3}(?:\.\d{3})*(?:,\d+)?");
                                if (m.Success) dt.TTien = double.Parse(m.Value.Replace(".", ""));
                            }

                            // Fuzzy Match & Cache
                            string keyCache = NormalizeVietnameseString(dt.Ten);
                            if (cacheMatHangTrongHoaDon.TryGetValue(keyCache, out var cached))
                            {
                                dt.SoHieu = cached.SoHieu; dt.Percent = cached.Percent;
                            }
                            else
                            {
                                Xulysohieuvattu(dt);
                                cacheMatHangTrongHoaDon[keyCache] = dt;
                            }

                            if (type == 1 && (tchat == 3 || daGiam)) dt.TKCo = "711";
                            dt.Ten = Helpers.ConvertUnicodeToVni(dt.Ten);
                            dt.Percent = Math.Round(dt.Percent);
                            tbImport.tbImportDetails.Add(dt);
                        }
                        catch (Exception ex)
                        {
                            XtraMessageBox.Show(ten);
                        }

                    }
                    //Tiến hành làm tròn và phân bổ thằng cuối cùng
                    foreach (var lt in tbImport.tbImportDetails)
                    {
                        lt.TTien = Math.Round(lt.TTien);
                    }
                    double sodu = Math.Round(finalTotal) - tbImport.tbImportDetails.Sum(m => m.TTien);
                    if (sodu > 0 && sodu <= 1)
                    {
                        tbImport.tbImportDetails.LastOrDefault().TTien += sodu;
                    }

                    //Thưc hiện thiếu tiền so với trc thuế
                    if (tbImport.TgTCThue != finalTotal)
                    {
                        sodu = tbImport.TgTCThue - finalTotal;
                        if (sodu > 0 && sodu <= 1)
                        {
                            tbImport.tbImportDetails.LastOrDefault().TTien += sodu;
                            if (tbImport.TgTCThue1 + sodu == tbImport.TgTCThue)
                            {
                                tbImport.TgTCThue1 += sodu;
                            }
                        }
                    }
                    // 12. Hoàn tất & Lưu
                    if (string.IsNullOrEmpty(tbImport.Noidung) && tbImport.tbImportDetails.Count > 0)
                        tbImport.Noidung = tbImport.tbImportDetails[0].Ten;

                    //Thông tin thay thế
                    string SHDCLQuan = ttChung.SelectSingleNode("//SHDCLQuan")?.InnerText;
                    if (!string.IsNullOrEmpty(SHDCLQuan))
                    {
                        try
                        {
                            DateTime NLHDCLQuan = DateTime.Parse(ttChung.SelectSingleNode("//NLHDCLQuan")?.InnerText);
                            string KHHDCLQuan = ttChung.SelectSingleNode("//KHHDCLQuan")?.InnerText;
                            tbImport.Noidung = Helpers.ConvertUnicodeToVni($"Thay thế cho ký hiệu hóa đơn {KHHDCLQuan}, số hóa đơn {SHDCLQuan}, ngày lập {NLHDCLQuan.ToShortDateString()}");
                        }
                        catch (Exception ex)
                        {

                        }

                    }
                    if (type == 1) { Xuly711(tbImport); }
                    else { /*Xuly5211(tbImport);*/ }

                    //await SaveDataXmlOne(tbImport, type);
                    stt++;
                }
            }
            catch (Exception ex) { XtraMessageBox.Show($"Lỗi xử lý file {pathXml}: {ex.Message}"); }
            finally { sothutu++; }
            //Trước khi import thêm vào cho lookupTbImport
            var keys = NormalizeTbImportKey(tbImport.Mst, tbImport.SHDon, tbImport.NLap, type);
            lookupTbImport.Add(keys);
            if (isAddhd == true)
                return tbImport;
            else
                return null;
        }
        private string NormalizeNameForSearch(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // 1. Chuẩn hóa dấu tiếng Việt (nếu có hàm Helpers.NormalizeVietnameseString)
            string result = Helpers.NormalizeVietnameseString(input);

            // 2. Chuyển về chữ thường
            result = result.ToLower().Trim();

            // 3. Thay thế các từ đồng nghĩa
            if (_synonymDictionary != null)
            {
                foreach (var synonym in _synonymDictionary)
                {
                    // Thay thế từ khóa
                    if (result.Contains(synonym.Key))
                    {
                        result = result.Replace(synonym.Key, synonym.Value);
                    }
                }
            }

            // 4. Xóa các ký tự đặc biệt thừa
            result = Regex.Replace(result, @"[^\w\s]", " ");

            // 5. Xóa khoảng trắng thừa
            result = Regex.Replace(result, @"\s+", " ").Trim();

            return result;
        }
        private Dictionary<string, string> _synonymDictionary;
        Regex regex = new Regex(@"(\d+(g|ml|L|kg)|x\d+|(\d+\s*cái))", RegexOptions.IgnoreCase);
        // Hàm mở rộng: Thêm từ đồng nghĩa động
        public void AddSynonym(string original, string normalized)
        {
            if (_synonymDictionary == null)
                _synonymDictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (!_synonymDictionary.ContainsKey(original.ToLower()))
            {
                _synonymDictionary[original.ToLower()] = normalized.ToLower();
            }
        }

        // Hàm mở rộng: Thêm nhiều từ đồng nghĩa cùng lúc
        public void AddSynonyms(Dictionary<string, string> synonyms)
        {
            if (_synonymDictionary == null)
                _synonymDictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in synonyms)
            {
                if (!_synonymDictionary.ContainsKey(item.Key.ToLower()))
                {
                    _synonymDictionary[item.Key.ToLower()] = item.Value.ToLower();
                }
            }
        }
        private bool _isIndexBuilt = false;
        private Dictionary<string, HashSet<string>> _keywordIndex; // keyword -> list key
        private Dictionary<string, HashSet<string>> _quyCachIndex; // quyCach -> list key
        private void BuildIndexes()
        {
            if (_isIndexBuilt) return;

            _keywordIndex = new Dictionary<string, HashSet<string>>();
            _quyCachIndex = new Dictionary<string, HashSet<string>>();

            foreach (var kvp in _optimizedVatTu)
            {
                // Index theo từ khóa (chỉ lấy từ dài >= 3 ký tự)
                var words = kvp.Value.TenChuan.Split(new[] { ' ', '-', ',', ';', '/' }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(w => w.Length >= 3);

                foreach (var word in words)
                {
                    if (!_keywordIndex.ContainsKey(word))
                        _keywordIndex[word] = new HashSet<string>();
                    _keywordIndex[word].Add(kvp.Key);
                }

                // Index theo quy cách
                if (!string.IsNullOrEmpty(kvp.Value.QuyCach))
                {
                    if (!_quyCachIndex.ContainsKey(kvp.Value.QuyCach))
                        _quyCachIndex[kvp.Value.QuyCach] = new HashSet<string>();
                    _quyCachIndex[kvp.Value.QuyCach].Add(kvp.Key);
                }
            }

            _isIndexBuilt = true;
        }
        private void AddSynonymIfNeeded(string original, string normalized)
        {
            if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(normalized))
                return;

            if (!_synonymDictionary.ContainsKey(original.ToLower()))
            {
                _synonymDictionary[original.ToLower()] = normalized.ToLower();
            }
        }
        private void InitializeSynonymDictionary()
        {
            _synonymDictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            AddSynonymIfNeeded("sài gòn", "saigon");
            AddSynonymIfNeeded("sai gon", "saigon");
            AddSynonymIfNeeded("sài gòn", "saigon");

            // ✅ GIỮ LẠI: Thương hiệu
            AddSynonymIfNeeded("cocacola", "coca cola");
            AddSynonymIfNeeded("cô ca", "coca cola");
            AddSynonymIfNeeded("cô ca cô la", "coca cola");
            AddSynonymIfNeeded("pesi", "pepsi");
            AddSynonymIfNeeded("redbull", "red bull");
        }
        private string ExtractMainName(string fullName)
        {
            if (string.IsNullOrEmpty(fullName)) return fullName;

            // Bỏ phần trong ngoặc đơn, ngoặc vuông, ngoặc nhọn
            string result = Regex.Replace(fullName, @"\(.*?\)", "").Trim();
            result = Regex.Replace(result, @"\[.*?\]", "").Trim();
            result = Regex.Replace(result, @"\{.*?\}", "").Trim();

            // Xóa khoảng trắng thừa
            result = Regex.Replace(result, @"\s+", " ").Trim();

            return result;
        }
        private void Xulysohieuvattu(TbImportDetail tbImportDetail)
        {
            if (tbImportDetail == null || string.IsNullOrEmpty(tbImportDetail.Ten))
                return;

            if (!_isIndexBuilt) BuildIndexes();
            if (_synonymDictionary == null) InitializeSynonymDictionary();

            string originalTen = tbImportDetail.Ten?.Trim() ?? "";
            string normalizedTen = NormalizeNameForSearch(originalTen);
            string quyCach = regex.Match(normalizedTen).Value;
            string donViTinh = tbImportDetail.DVT?.Trim()?.ToLower() ?? "";

            Console.WriteLine($"🔍 Đang tìm: {normalizedTen}");
            Console.WriteLine($"   Quy cách: '{quyCach}'");

            double minPercent = 80;

            // ========== 1. TÌM CHÍNH XÁC ==========
            var exactMatch = _optimizedVatTu
                .FirstOrDefault(kvp =>
                    (NormalizeNameForSearch(kvp.Value.TenChuan) == normalizedTen ||
                     NormalizeNameForSearch(kvp.Value.TenPhuChuan) == normalizedTen) &&
                    (string.IsNullOrEmpty(quyCach) || kvp.Value.QuyCach == quyCach));

            if (!exactMatch.Equals(default(KeyValuePair<string, (string, string, string, string, double, double)>)))
            {
                tbImportDetail.SoHieu = exactMatch.Key;
                tbImportDetail.Percent = 100;
                tbImportDetail.DVT = exactMatch.Value.DonVi;
                Console.WriteLine($"✅ Tìm chính xác: {exactMatch.Value.TenChuan}");
                return;
            }

            // ========== 2. TÁCH TỪ KHÓA ==========
            var words = normalizedTen.Split(new[] { ' ', '-', ',', ';', '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(w => w.Length >= 3)
                .ToList();

            Console.WriteLine($"   Từ khóa: {string.Join(", ", words)}");

            var phrases = new List<string>();
            for (int i = 0; i < words.Count - 1; i++)
            {
                string phrase = words[i] + " " + words[i + 1];
                if (phrase.Length >= 5)
                {
                    phrases.Add(phrase);
                }
            }

            Console.WriteLine($"   Cụm từ: {string.Join(", ", phrases)}");

            // ========== 3. SÀNG LỌC ỨNG VIÊN ==========
            var candidateKeys = new HashSet<string>();

            foreach (var word in words)
            {
                if (_keywordIndex != null && _keywordIndex.ContainsKey(word))
                {
                    foreach (var key in _keywordIndex[word])
                    {
                        candidateKeys.Add(key);
                    }
                }
            }

            foreach (var phrase in phrases)
            {
                if (_keywordIndex != null && _keywordIndex.ContainsKey(phrase))
                {
                    foreach (var key in _keywordIndex[phrase])
                    {
                        candidateKeys.Add(key);
                    }
                }
            }

            if (!string.IsNullOrEmpty(quyCach) && _quyCachIndex != null && _quyCachIndex.ContainsKey(quyCach))
            {
                foreach (var key in _quyCachIndex[quyCach])
                {
                    candidateKeys.Add(key);
                }
            }

            Console.WriteLine($"   Số ứng viên tìm được: {candidateKeys.Count}");

            if (!candidateKeys.Any())
            {
                int count = 0;
                foreach (var kvp in _optimizedVatTu)
                {
                    if (count >= 100) break;
                    count++;
                    candidateKeys.Add(kvp.Key);
                }
                Console.WriteLine($"   Fallback: lấy 100 item đầu tiên");
            }

            // ========== 4. TÍNH ĐIỂM ==========
            var results = new List<(string Key, double Percent, string TenChuan, string QuyCach, string DonVi, int MatchCount)>();

            foreach (var key in candidateKeys)
            {
                if (_optimizedVatTu.TryGetValue(key, out var vatTu))
                {
                    string tenChuanHoa = NormalizeNameForSearch(vatTu.TenChuan);
                    string tenKhongNgoac = ExtractMainName(tenChuanHoa);
                    string tenHoaDonKhongNgoac = ExtractMainName(normalizedTen);

                    // So sánh tên không ngoặc
                    int tokenScoreNoBracket = Fuzz.TokenSetRatio(tenKhongNgoac, tenHoaDonKhongNgoac);
                    int partialScoreNoBracket = Fuzz.PartialRatio(tenKhongNgoac, tenHoaDonKhongNgoac);
                    double percentNoBracket = Math.Max(tokenScoreNoBracket, partialScoreNoBracket);

                    // So sánh tên đầy đủ
                    int tokenScore = Fuzz.TokenSetRatio(tenChuanHoa, normalizedTen);
                    int partialScore = Fuzz.PartialRatio(tenChuanHoa, normalizedTen);
                    double percent = Math.Max(tokenScore, partialScore);

                    // Lấy điểm cao nhất
                    double finalPercent = Math.Max(percent, percentNoBracket);

                    // Đếm số từ khóa khớp
                    int matchCount = 0;
                    foreach (var word in words)
                    {
                        if (tenChuanHoa.Contains(word) || tenKhongNgoac.Contains(word))
                            matchCount++;
                    }

                    finalPercent += matchCount * 5;

                    // So sánh quy cách
                    string quyCachTrongKho = vatTu.QuyCach?.ToLower()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(quyCach) && !string.IsNullOrEmpty(quyCachTrongKho))
                    {
                        if (quyCachTrongKho == quyCach || quyCachTrongKho.Contains(quyCach))
                        {
                            finalPercent += 20;
                        }
                    }

                    // ========== GIỚI HẠN ĐIỂM ==========
                    // Kiểm tra xem có phải đang khớp qua tên không ngoặc không
                    bool isMatchByNoBracket = percentNoBracket > 80 && percentNoBracket > percent;

                    if (isMatchByNoBracket)
                    {
                        // Nếu khớp qua tên không ngoặc (thiếu thông tin trong ngoặc)
                        // Giới hạn tối đa 90%
                        if (finalPercent > 90)
                        {
                            finalPercent = 90;
                        }
                    }
                    else if (percent > 80)
                    {
                        // Nếu khớp qua tên đầy đủ nhưng không chính xác 100%
                        // Giới hạn tối đa 95%
                        if (finalPercent > 95)
                        {
                            finalPercent = 95;
                        }
                    }

                    // Debug
                    if (vatTu.TenChuan.Contains("Tiger") || vatTu.TenChuan.Contains("tiger"))
                    {
                        Console.WriteLine($"   Kiểm tra: {vatTu.TenChuan}");
                        Console.WriteLine($"      Điểm: {finalPercent}%");
                        Console.WriteLine($"      Không ngoặc: {percentNoBracket}%");
                        Console.WriteLine($"      Từ khớp: {matchCount}");
                        Console.WriteLine($"      Quy cách: '{quyCachTrongKho}'");
                        Console.WriteLine($"      Khớp không ngoặc: {isMatchByNoBracket}");
                    }

                    if (finalPercent >= minPercent)
                    {
                        results.Add((key, Math.Min(finalPercent, 100), vatTu.TenChuan, vatTu.QuyCach, vatTu.DonVi, matchCount));
                    }
                }
            }

            // ========== 5. CHỌN KẾT QUẢ ==========
            if (results.Any())
            {
                var sorted = results
                    .OrderByDescending(x => x.MatchCount)
                    .ThenByDescending(x => x.Percent)
                    .ToList();

                var best = sorted.First();
                tbImportDetail.SoHieu = best.Key;
                tbImportDetail.Percent = best.Percent;
                tbImportDetail.DVT = best.DonVi;

                Console.WriteLine($"✅ Tìm thấy: {best.TenChuan}");
                Console.WriteLine($"   Độ tương đồng: {best.Percent}%");
                Console.WriteLine($"   Quy cách: {best.QuyCach}");

                if (sorted.Count > 1)
                {
                    Console.WriteLine($"   📋 Các kết quả khác:");
                    foreach (var item in sorted.Skip(1).Take(3))
                    {
                        Console.WriteLine($"      - {item.TenChuan} (Điểm: {item.Percent}%)");
                    }
                }
            }
            else
            {
                tbImportDetail.SoHieu = GenerateResultString(Helpers.NormalizeVietnameseString(normalizedTen));
                tbImportDetail.Percent = 0;
                Console.WriteLine($"❌ Không tìm thấy vật tư cho: {normalizedTen}");

            }
        }
        private static string RemoveVietnameseDiacritics(string str)
        {
            // Mảng chứa ký tự có dấu
            str = str.ToLower();
            str = Regex.Replace(str, "[àáạảãâầấậẩẫăằắặẳẵ]", "a");
            str = Regex.Replace(str, "[èéẹẻẽêềếệểễ]", "e");
            str = Regex.Replace(str, "[ìíịỉĩ]", "i");
            str = Regex.Replace(str, "[òóọỏõôồốộổỗơờớợởỡ]", "o");
            str = Regex.Replace(str, "[ùúụủũưừứựửữ]", "u");
            str = Regex.Replace(str, "[ỳýỵỷỹ]", "y");
            str = Regex.Replace(str, "đ", "d");

            // Thay thế khoảng trắng bằng dấu gạch ngang
            //  str = Regex.Replace(str, " ", "-");
            str = str.Replace(",", "");
            str = str.Replace(".", "");
            str = str.Replace("*", "x");
            // Thay thế tất cả các âm "o" có dấu thành "o" không dấu
            str = str.Replace("ó", "o");
            str = str.Replace("ò", "o");
            str = str.Replace("õ", "o");
            str = str.Replace("ọ", "o");
            str = str.Replace("ỏ", "o");
            str = str.Replace("ô", "o");
            str = str.Replace("ơ", "o");
            str = str.Replace("'", "");
            return str;
        }
        public static string GenerateResultString(string input)
        {
            // Tìm từ đầu tiên (không cần loại bỏ dấu toàn bộ)
            string firstWord = input.Split(' ')[0];

            // Loại bỏ dấu tiếng Việt cho từ đầu tiên
            string normalizedFirstWord = RemoveVietnameseDiacritics(firstWord).Replace("á", "a");
            if (normalizedFirstWord.Length >= 10)
            {
                if (normalizedFirstWord.Length > 10)
                {
                    normalizedFirstWord = normalizedFirstWord.Substring(0, 10);
                }
                normalizedFirstWord = char.ToUpper(normalizedFirstWord[0]) + normalizedFirstWord.Substring(1);
            }
            // Tạo 4 số ngẫu nhiên từ 1 đến 9
            string randomNumbers = GenerateRandomNumbers(4);

            // Kết hợp từ đầu tiên với 4 số ngẫu nhiên
            return CapitalizeFirstLetter(normalizedFirstWord).ToUpper() + randomNumbers;
        }
        static string CapitalizeFirstLetter(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input; // Kiểm tra chuỗi rỗng hoặc null

            return char.ToUpper(input[0]) + input.Substring(1);
        }
        private static Random random = new Random(); // Tạo Random tĩnh để tái sử dụng
        private static string GenerateRandomNumbers(int length)
        {
            string randomNumbers = "";
            HashSet<int> generatedNumbers = new HashSet<int>(); // Sử dụng HashSet để lưu các số đã tạo

            while (randomNumbers.Length < length)
            {
                // Sinh số ngẫu nhiên từ 1 đến 9
                int number = random.Next(1, 10);

                // Kiểm tra nếu số đó chưa được tạo
                if (!generatedNumbers.Contains(number))
                {
                    randomNumbers += number.ToString();
                    generatedNumbers.Add(number); // Thêm số vào HashSet
                }
            }

            return randomNumbers;
        }
        private bool CheckExistKH(string mst)
        {

            //Nếu có Mã s61 thuế
            if (!string.IsNullOrEmpty(mst))
            {
                if (tbKhachhang.AsEnumerable().Any(row => row.Field<string>("MST") == mst || row.Field<string>("SoHieu") == mst))
                {
                    return true;
                }
            }

            return false;
        }
        public string GenerateAbbreviation(string fullName, List<string> existingNames)
        {
            // Tách tên thành từng phần
            string[] nameParts = fullName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string abbreviation = "";

            // Tạo viết tắt
            foreach (string part in nameParts)
            {
                abbreviation += part[0].ToString().ToLower();
            }

            // Kiểm tra sự tồn tại của viết tắt
            int counter = 1;
            string uniqueAbbreviation = abbreviation.ToUpper();

            while (existingNames.Contains(uniqueAbbreviation))
            {
                uniqueAbbreviation = abbreviation + "-" + counter;
                counter++;
            }

            return uniqueAbbreviation;
        }
        public void InitCustomer(int Maphanloai, string Sohieu, string Ten, string Diachi, string Mst, string cccd, string sdt)
        {
            if (string.IsNullOrEmpty(sdt))
                sdt = "xxx";
            int randNumber = 0;
            Random random = new Random();

            //Xử lý địa chỉ
            string diachiKHVni = !string.IsNullOrEmpty(Diachi) ? Helpers.ConvertUnicodeToVni(Diachi) : Helpers.ConvertUnicodeToVni("Bổ sung địa chỉ");

            if (string.IsNullOrEmpty(Mst))
            {
                //Truong hợp ko có mst và cccd
                if (string.IsNullOrEmpty(cccd))
                {

                    Sohieu = GenerateAbbreviation(Helpers.ConvertVniToUnicode(Ten), tbKhachhang.AsEnumerable().Select(row => row.Field<string>("SoHieu")).ToList()).ToUpper();
                    csohieu = Sohieu;
                    Mst = "00";

                    //Xử lý khi số hiệu bị trùng
                    int suffix = 1;
                    string originalSohieu = Sohieu;

                    while (tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                    {
                        Sohieu = $"{originalSohieu}_{suffix}";
                        suffix++;
                    }
                }
                //Không có mst nhưng có cccd
                else
                {
                    Sohieu = cccd.Substring(cccd.Length - 6);
                    Mst = cccd;
                }
            }
            else
            {
                if (string.IsNullOrEmpty(Sohieu))
                {
                    Sohieu = Helpers.GetLastFourDigits(Mst.Replace("-", ""));

                    string tenKHVni = Helpers.ConvertUnicodeToVni(Ten);

                    //Xử lý khi số hiệu bị trùng
                    if (tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                    {
                        Sohieu = "0" + Sohieu;
                    }
                    if (tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                    {
                        Sohieu = "00" + Sohieu;
                    }
                }
            }
            //Nếu tồn tại so hiệu r, sẽ thêm kí tự
            if (tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
            {
                Sohieu = Sohieu + "_1";
            }
            if (Mst == Sohieu && Mst.Length <= 8)
            {
                Mst = "00";
            }
            else
            {
                if (Mst.Length > 8)
                {
                    Sohieu = Helpers.GetLastFourDigits(Mst.Replace("-", ""));
                    //Kiểm tra SoHieu co trung thêm 1 lần
                    if (tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                    {
                        Sohieu = "0" + Sohieu;
                    }
                    if (tbKhachhang.AsEnumerable().Any(row => row.Field<string>("SoHieu") == Sohieu))
                    {
                        Sohieu = "00" + Sohieu;
                    }
                }
            }

            if (mstcongty == "3501972322")
            {
                int lastid = 0;
                string qr = "SELECT MAX(MaSo) FROM KhachHang";
                var getlastid = ExecuteQuery(qr).Rows[0]["Expr1000"].ToString();
                string query = @"
        INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu,Ten,DiaChi,MST,Tel)
        VALUES (?,?,?,?,?,?,?)";


                // Khai báo mảng tham số với đủ 10 tham số
                OleDbParameter[] parameters = new OleDbParameter[]
                {
               new OleDbParameter("?", int.Parse(getlastid)+1),
               new OleDbParameter("?", Maphanloai),
               new OleDbParameter("?", Sohieu),
               new OleDbParameter("?", Ten),
               new OleDbParameter("?", diachiKHVni),
               new OleDbParameter("?", Mst),
               new OleDbParameter("?", sdt),
                };

                // Thực thi truy vấn và lấy kết quả
                try
                {
                    int a = ExecuteQueryResult(query, parameters);
                    query = "SELECT * FROM KhachHang"; // Giả sử bạn muốn lấy tất cả dữ liệu từ bảng KhachHang
                    tbKhachhang = ExecuteQuery(query);
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message + "    " + Ten + "   " + cccd);
                }
            }
            else
            {

                string query = @"
        INSERT INTO KhachHang (MaPhanLoai,SoHieu,Ten,DiaChi,MST,Tel)
        VALUES (?,?,?,?,?,?)";


                // Khai báo mảng tham số với đủ 10 tham số
                OleDbParameter[] parameters = new OleDbParameter[]
                {
               new OleDbParameter("?", Maphanloai),
               new OleDbParameter("?", Sohieu),
               new OleDbParameter("?", Ten),
               new OleDbParameter("?", diachiKHVni),
               new OleDbParameter("?", Mst),
               new OleDbParameter("?", sdt),
                };

                // Thực thi truy vấn và lấy kết quả
                try
                {
                    int a = ExecuteQueryResult(query, parameters);
                    query = "SELECT * FROM KhachHang"; // Giả sử bạn muốn lấy tất cả dữ liệu từ bảng KhachHang
                    tbKhachhang = ExecuteQuery(query);
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message + "    " + Ten + "   " + cccd);
                }
            }
        }
        private void GetMST()
        {
            string query = "SELECT * FROM License";

            // Tạo mảng tham số với giá trị cho câu lệnh SQL

            var kq = ExecuteQuery(query, null);
            if (kq.Rows.Count > 0)
            {
                MSTCongTY = kq.Rows[0]["MaSoThue"].ToString();
                CCCD = kq.Rows[0]["CCCD"].ToString();
            }
        }
        private string MSTCongTY = "";
        private string CCCD = "";
        private Dictionary<string, TbImportDetail> cacheMatHangTrongHoaDon;
        private bool Loaiborow(string name)
        {
            if (name.ToLower().Contains("điều chỉnh"))
                return true;
            return false;
        }
        string CapitalizeFirstLetters(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input; // Kiểm tra chuỗi rỗng hoặc null

            return char.ToUpper(input[0]) + input.Substring(1);
        }
          string csohieu = "";
        private bool Kiemtrahoadon(string SHDon, DateTime NLap, string MST, int type)
        {
            // Tạo tuple từ 3 tham số
            var key = (MST, SHDon, NLap, type);

            // Kiểm tra trong lookup
            return lookupTbImport.Contains(key);
        }
        private bool KiemtrahoadonCT(string SoHD, string KyHieu, DateTime NLap, string Mst, int tpye)
        {
            if (Mst == "KL")
                Mst = "00";
            if (Mst.Length < 10)
                return lookupHoaDonCT.Any(m => m.SoHD == SoHD && m.KyHieu == KyHieu && m.NLap == NLap && m.Type == type);
            return lookupHoaDonCT.Contains((Mst, SoHD, KyHieu, NLap, tpye));
        }
        int type = 0;
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

                public FileImportDetail(int id, string ten, int parentId, string soHieu, double soluong, double dongia, string dVT, string maCT, string tkNo, string tkCo, double ttien, string percent, int tchat, int vat)
                {
                    Ten = ten;
                    ParentId = parentId;
                    SoHieu = soHieu;
                    Soluong = soluong;
                    Dongia = dongia;
                    DVT = dVT;
                    MaCT = maCT;
                    TKNo = tkNo;
                    TKCo = tkCo;
                    TTien = ttien;
                    ID = id;
                    Percent = percent;
                    Tchat = tchat;
                    IsMacdinh = 0;
                    SoPSGoc = ttien;
                    VAT = vat;
                }
            }
            public static int Id = 1;
            public FileImport(bool chk, string path, string shdon, string khhdon, DateTime nlap, DateTime ntao, string ten, string noidung, string tkno, string tkco, int tkthue, string mst, double tongTien, int vat, int type, string tenTP, bool isacess, double tPhi, double tgTCThue, double tgTThue, bool _isHaschild, int _InvoiceType, double tvat, int vat2, double tvat2, int vat3, double tvat3, string macdinhstatus, bool khautruthue)
            {
                ID = Id;
                SHDon = shdon;
                KHHDon = khhdon;
                NLap = nlap;
                Ngaytao = ntao;
                Ten = ten;
                Noidung = noidung;
                TKCo = tkco;
                TKNo = tkno;
                TkThue = tkthue;
                Mst = mst;
                TongTien = tongTien;
                Vat = vat;
                Id += 1;
                fileImportDetails = new List<FileImportDetail>();
                Type = type;
                Checked = chk;
                Path = path;
                SoHieuTP = tenTP;
                isAcess = isacess;
                TPhi = tPhi;
                TgTCThue = tgTCThue;
                TgTThue = tgTThue;
                isHaschild = _isHaschild;
                InvoiceType = _InvoiceType;
                TVat = tvat;
                Vat2 = vat2;
                TVat2 = tvat2;
                Vat3 = vat3;
                TVat3 = tvat3;
                Macdinhstatus = macdinhstatus;
                Khautruthue = khautruthue;
            }



        }
        private BindingList<FileImport> lstImportVao = new BindingList<FileImport>();
        private BindingList<FileImport> lstImportRa = new BindingList<FileImport>();
        public static double SafeParse(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return 0;
            // Thay thế dấu phẩy thành dấu chấm để đồng nhất định dạng chuẩn C#
            value = value.Replace(",", ".");
            if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            return 0;
        }
        DataTable tbDinhDanhtaikhoan = new DataTable();
        public DataTable tbKhachhang = new DataTable();
        public int sothutu = 1;
        bool isAddhd = true;
        private (string MST, string SHDon, DateTime NLap, int Types) NormalizeTbImportKey(
string mst, string shDon, DateTime nLap, int Types)
        {
            return (
                (mst ?? "").Trim(),
                Helpers.RemoveLeadingZeros(shDon ?? "").Trim(),
                nLap.Date,
                Types
            );
        }
        public static string NormalizeVietnameseString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "";
            input = input.Replace("*", "x");
            input = input.Replace(" lít", "l");
            input = input.Replace("( ", "(");
            // 4. Bỏ nội dung trong ngoặc
            // input = Regex.Replace(input, @"\(.*?\)", "");
            // input= Regex.Replace(input, @"\*(\d+)", "x$1");

            // 7. Tách số và chữ (455ml → 455 ml)
            // input = Regex.Replace(input, @"(\d+)([a-zA-Z])", "$1 $2");
            // input = Regex.Replace(input, @"([a-zA-Z])(\d+)", "$1 $2");
            //input = Regex.Replace(input, @"(\d+)\s*x\s*(\d+)\s*ml", "$2 ml x $1");

            //input = RemoveLeadingSpecialCharacters(input);
            //Bỏ đi tab
            input = input.Replace("\t", ""); // Thay thế ký tự tab bằng chuỗi rỗng
            input = input.Normalize(NormalizationForm.FormC);

            return input.Trim();
        }
        private async Task SaveAllInvoicesBulk(List<TbImport> invoices, int type,string connectionString)
        {
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                await conn.OpenAsync();
                using (OleDbTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        string sqlParent = @"INSERT INTO tbImport (SHDon, KHHDon, NLap, Ten, Noidung, TKNo, TKCo, TkThue, Mst, [Status], Ngaytao, TongTien, Vat, TPhi, TgTCThue, TgTThue, [Type], InvoiceType, IsHaschild, TVat, Vat2, TVat2, Vat3, TVat3, TgTCThue1, TgTCThue2, TgTCThue3, Khmshdon, hdon, [Path]) 
                                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

                        string sqlDetail = @"INSERT INTO tbimportdetail (ParentId, SoHieu, SoLuong, DonGia, DVT, Ten, MaCT, TKNo, TKCo, TTien, [Percent], Tchat,SoPSGoc,VAT) 
                                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

                        foreach (var item in invoices)
                        {
                            int parentID = 0;

                            // 1. Lưu Hóa đơn chính (Parent)
                            using (OleDbCommand cmdParent = new OleDbCommand(sqlParent, conn, trans))
                            {
                                cmdParent.Parameters.AddWithValue("?", item.SHDon ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.KHHDon ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.NLap);
                                cmdParent.Parameters.AddWithValue("?", item.Ten ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.Noidung ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.TKNo ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.TKCo ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.TkThue ?? "");
                                cmdParent.Parameters.AddWithValue("?", item.Mst ?? "");
                                cmdParent.Parameters.AddWithValue("?", "0"); // Status
                                cmdParent.Parameters.AddWithValue("?", DateTime.Now.ToShortDateString());
                                cmdParent.Parameters.AddWithValue("?", item.TongTien);
                                cmdParent.Parameters.AddWithValue("?", item.Vat);
                                cmdParent.Parameters.AddWithValue("?", item.TPhi ?? "0");
                                cmdParent.Parameters.AddWithValue("?", Math.Round(item.TgTCThue));
                                cmdParent.Parameters.AddWithValue("?", Math.Round(item.TgTThue));
                                cmdParent.Parameters.AddWithValue("?", type);
                                cmdParent.Parameters.AddWithValue("?", "0"); // InvoiceType
                                cmdParent.Parameters.AddWithValue("?", "1"); // IsHaschild
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

                                await cmdParent.ExecuteNonQueryAsync();
                            }

                            // 2. Lấy ID tự tăng vừa tạo
                            using (OleDbCommand cmdId = new OleDbCommand("SELECT @@IDENTITY", conn, trans))
                            {
                                var objId = await cmdId.ExecuteScalarAsync();
                                parentID = Convert.ToInt32(objId);
                            }

                            // 3. Lưu chi tiết hàng hóa (Details)
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
                                    cmdDetail.Parameters.AddWithValue("?", ""); // MaCT
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

                        // Chốt giao dịch: Ghi toàn bộ xuống ổ cứng
                        trans.Commit();
                        this.Close();
                        //XtraMessageBox.Show($"Đã lưu thành công tổng cộng {invoices.Count} hóa đơn vào Database!", "Thông báo");
                    }
                    catch (Exception ex)
                    {
                        // Nếu có bất kỳ lỗi nào, hủy bỏ toàn bộ để tránh dữ liệu rác
                        trans.Rollback();
                        XtraMessageBox.Show("Lỗi hệ thống khi lưu hàng loạt: " + ex.Message, "Lỗi Database");
                    }
                }
            }
        }
        public void Xuly711(TbImport fileImport)
        {
            var getlist711 = fileImport.tbImportDetails.Where(m => m.TKCo == "711").ToList();

            if (getlist711.Count > 0 && tbLicense.Rows[0].Field<string>("col711") == "1")
            {
                List<TbImportDetail> listdathuchien = new List<TbImportDetail>();

                //Trường hợp chỉ có 1 dòng 711
                //Lấy ra dòng 711;
                if (getlist711.Count == 1)
                {
                    var get711 = getlist711.Where(m => m.TKCo == "711").FirstOrDefault();
                    if (get711 != null)
                    {
                        string pattern = @"\d+"; // Tìm một hoặc nhiều chữ số

                        Match match = Regex.Match(get711.Ten, pattern);
                        //Kiển tra xem có phải chiết khấu có % không

                        //Tính lại giá tiền cho các dòng
                        var remainlist = fileImport.tbImportDetails.Where(m => m != get711 && !string.IsNullOrEmpty(m.DVT) && m.Dongia != 0).ToList();
                        int index = 0;
                        double sumtotal = 0;

                        double totalth = remainlist.Sum(m => m.TTien);
                        double total711 = get711.TTien;
                        foreach (var it2 in remainlist)
                        {
                            if (index < remainlist.Count - 1)
                            {
                                it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                            }
                            else
                            {
                                it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                                double reTotal = Math.Round(fileImport.TgTCThue - remainlist.Sum(m => m.TTien));
                                if (reTotal > 0)
                                    it2.TTien += reTotal;
                                if (reTotal == -1)
                                    it2.TTien -= 1;
                            }
                            //Cập nhật vô database

                            index += 1;
                        }
                        //Xoá đi dòng 711
                        fileImport.tbImportDetails.Remove(get711);
                    }

                }
                //Trường hợp có nhiều 711
                else
                {
                    var get771s = getlist711.Where(m => m.TKCo == "711").ToList();
                    int finddata = 0;
                    foreach (var i7 in get771s)
                    {
                        //string pattern = @"(\d+[,.]?\d*)%|chiết khấu\s*(\d+)";


                        //var match = Regex.Match(i7.Ten, pattern);
                        //double percent = 0;
                        //if (match.Success)
                        //{
                        //    string soChietKhau = match.Groups[1].Value;
                        //    if (string.IsNullOrEmpty(soChietKhau))
                        //    {
                        //        soChietKhau = match.Groups[2].Value;
                        //    }
                        //    soChietKhau = soChietKhau.Replace(",", ".");
                        //    percent = double.Parse(soChietKhau);
                        //    foreach (var ftt in fileImport.tbImportDetails.Where(m => m.TKCo != "711"))
                        //    {
                        //        var sodu = Math.Round(ftt.TTien * percent / 100) - i7.TTien;
                        //        if (sodu >= 0 && sodu <= 1)
                        //        {
                        //            finddata += 1;
                        //            ftt.TTien = ftt.TTien - i7.TTien;
                        //            listdathuchien.Add(i7);
                        //        }
                        //    }

                        //}
                        fileImport.tbImportDetails.Remove(i7);
                    }
                    //Trường hợp ko tìm dc thì buộc phải tìm % để phân bổ



                    //if (finddata != get771s.Count)
                    if (1 < 2)
                    {

                        //Tính lại giá tiền cho các dòng
                        var remainlist = fileImport.tbImportDetails.Where(m => m.TKCo != "711" && !string.IsNullOrEmpty(m.DVT) && m.Dongia != 0).ToList();
                        double totalth = remainlist.Sum(m => m.TTien);
                        double total711 = get771s.Sum(m => m.TTien);
                        if (total711 < 0)
                        {
                            total711 = -total711;
                        }
                        int index = 0;
                        foreach (var it2 in remainlist)
                        {
                            if (remainlist.Count > 1)
                            {
                                if (index < remainlist.Count - 1)
                                {
                                    it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                                }
                                else
                                {
                                    it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                                    double reTotal = Math.Round(fileImport.TgTCThue - remainlist.Sum(m => m.TTien));
                                    it2.TTien += reTotal;
                                }
                            }
                            else
                            {
                                it2.TTien = Math.Round(it2.TTien - (total711 * it2.TTien / totalth));
                            }
                            index += 1;
                        }

                    }
                }

            }
            else
            {
                if (getlist711.Count > 1)
                {
                    var getfirst = getlist711.FirstOrDefault();
                    var lstremain = getlist711.Skip(1).ToList();
                    //Cập nhật lại tổng tiền cho first
                    getfirst.TTien = getfirst.TTien + lstremain.Sum(m => m.TTien);
                    if (getfirst.TTien < 0)
                    {
                        getfirst.TTien = -getfirst.TTien;
                    }
                    //Xoá các dòng thừa
                    foreach (var it in lstremain)
                    {
                        fileImport.tbImportDetails.Remove(it);
                    }
                }
            }
        }

        #endregion

        private async void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;

            try
            {
                int totalRuns = 3; // Số lần lặp lại

                for (int runCount = 1; runCount <= totalRuns; runCount++)
                {
                    //Đọc file xml
                    foreach (DataRow item in tbCompany.Rows)
                    {
                        string vbdbpath = item["Dbpath"]?.ToString() ?? "";
                        string connectionString2 = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                    "Data Source=" + vbdbpath + ";" +
                                    "Jet OLEDB:Database Password=1@35^7*9)1;";
                        LoadHoadonCT(connectionString2);
                        Loadtbimport(connectionString2);
                        string qrkh = "SELECT * FROM KhachHang"; // Giả sử bạn muốn lấy tất cả dữ liệu từ bảng KhachHang
                        tbKhachhang = ExecuteQuery2(qrkh, connectionString2);
                        XulylietkeHoaDon(2, connectionString2);
                    }
                
                    return;
                    Log($"🔄 ===== BẮT ĐẦU LẦN CHẠY {runCount}/{totalRuns} =====");
                    labelControl3.Text = $"🔄 ===== BẮT ĐẦU LẦN CHẠY {runCount}/{totalRuns} =====";
                    // Refresh lại danh sách công ty mỗi lần chạy
                    string query = @"SELECT * FROM tbCompany WHERE IsRun = 1 and Saoviet = ? order by STT";
                    string computerName = Environment.MachineName;

                    tbCompany = ExecuteQuery(query, new OleDbParameter("?", computerName));

                    if (tbCompany.Rows.Count == 0)
                    {
                        Log($"⚠️ Lần chạy {runCount}: Không có công ty nào đang hoạt động!");
                        continue;
                    }

                    Log($"🚀 Lần chạy {runCount}: Xử lý {tbCompany.Rows.Count} công ty (tối đa {txtSoluongtai.Text} cùng lúc)");

                    int maxParallel = int.Parse(txtSoluongtai.Text);
                    SemaphoreSlim semaphore = new SemaphoreSlim(maxParallel);
                    List<Task> tasks = new List<Task>(); 
                    foreach (DataRow item in tbCompany.Rows)
                    {
                        string vbdbpath = item["Dbpath"]?.ToString() ?? "";
                        string companyName = item["Name"]?.ToString() ?? "Unknown";

                        if (string.IsNullOrEmpty(vbdbpath))
                        {
                            Log($"⚠️ {companyName}: Không có Dbpath, bỏ qua!");
                            continue;
                        }

                        await semaphore.WaitAsync();

                        // Copy DataRow để tránh lỗi cross-thread
                        DataRow rowCopy = item;

                        tasks.Add(Task.Run(async () =>
                        {
                            try
                            {
                                // Cập nhật status trên UI thread
                                UpdateStatusOnUI(rowCopy, $"🔄 Lần {runCount}/{totalRuns} - Đang xử lý...");

                                await TaihoadonCongty(vbdbpath, rowCopy);

                                // Cập nhật status sau khi hoàn thành
                                UpdateStatusOnUI(rowCopy, $"✅ Lần {runCount}/{totalRuns} - Hoàn thành");
                                //Tiến hành đọc hoá đơn

                            }
                            catch (Exception ex)
                            {
                                Log($"❌ Lỗi {companyName} (Lần {runCount}): {ex.Message}");
                                UpdateStatusOnUI(rowCopy, $"❌ Lần {runCount}: {ex.Message}");
                            }
                            finally
                            {
                                semaphore.Release();
                                await Task.Delay(2000);
                            }
                        }));
                    }

                    await Task.WhenAll(tasks);
                    Log($"✅ Lần chạy {runCount}/{totalRuns} hoàn thành!");
                   
                    // Chờ giữa các lần chạy
                    if (runCount < totalRuns)
                    {
                        Log($"⏳ Chờ 2s trước lần chạy tiếp theo...");
                        await Task.Delay(2000);
                    }
                }

                Log($"✅ ===== HOÀN THÀNH TẤT CẢ {totalRuns} LẦN CHẠY =====");
            }
            catch (Exception ex)
            {
                Log($"❌ Lỗi: {ex.Message}");
            }
            finally
            {
                btnRun.Enabled = true;
            }
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
    #endregion
}