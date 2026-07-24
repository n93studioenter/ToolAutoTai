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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static ToolTaiHD.Form1;

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
            LoadHoadonCT();
            Loadtbimport();


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
        private void LoadHoadonCT()
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

                var data = ExecuteQuery(query);
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
        private void Loadtbimport()
        {
            try
            {
                string query = "SELECT * FROM tbimport";
                tbimports = ExecuteQuery(query);
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
        private async void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;

            try
            {
                int totalRuns = 5; // Số lần lặp lại

                for (int runCount = 1; runCount <= totalRuns; runCount++)
                {
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