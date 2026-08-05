using DevExpress.LookAndFeel;
using DevExpress.Skins;
using DevExpress.UserSkins;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ToolTaiHD
{
    internal static class Program
    {
        #region Win32 API
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool FlashWindow(IntPtr hWnd, bool bInvert);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // ✅ Tên Mutex duy nhất cho ứng dụng
            const string appMutexName = "ToolTaiHD_SingleInstance_Mutex_2024";

            bool createdNew;
            using (Mutex mutex = new Mutex(true, appMutexName, out createdNew))
            {
                if (!createdNew)
                {
                    // ✅ Đã có instance khác đang chạy → Kích hoạt instance cũ
                    ActivateExistingInstance();
                    return; // Thoát instance hiện tại
                }

                // ✅ Nếu là instance đầu tiên, chạy bình thường
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = (a, b, c, d) => true;

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // ✅ Kiểm tra AutoStart
                bool isAutoStart = args.Any(a => a.Equals("-autostart", StringComparison.OrdinalIgnoreCase));

                // ✅ Kiểm tra Task Scheduler
                bool isTaskScheduler = IsRunningFromTaskScheduler();

                // ✅ Tạo form với tham số isAutoMode
                Form1 form = new Form1(isAutoStart || isTaskScheduler);
               // Form2 form = new Form2();

                Application.Run(form);
            }
        }

        /// <summary>
        /// Kích hoạt instance đang chạy (đưa lên foreground, flash window)
        /// </summary>
        private static void ActivateExistingInstance()
        {
            try
            {
                // Lấy process hiện tại
                Process currentProcess = Process.GetCurrentProcess();
                string processName = currentProcess.ProcessName;

                // Tìm các process cùng tên (trừ process hiện tại)
                Process[] processes = Process.GetProcessesByName(processName);

                foreach (Process process in processes)
                {
                    // Bỏ qua process hiện tại
                    if (process.Id == currentProcess.Id)
                        continue;

                    IntPtr handle = process.MainWindowHandle;

                    if (handle != IntPtr.Zero)
                    {
                        // ✅ Nếu cửa sổ đang thu nhỏ → Restore
                        if (IsIconic(handle))
                        {
                            ShowWindow(handle, SW_RESTORE);
                        }

                        // ✅ Đưa lên foreground
                        SetForegroundWindow(handle);
                        BringWindowToTop(handle);

                        // ✅ Flash window để thu hút sự chú ý
                        FlashWindow(handle, true);

                        // ✅ Đợi một chút để window được active
                        Thread.Sleep(100);

                        // ✅ Thoát instance hiện tại
                        Environment.Exit(0);
                        return;
                    }
                }

                // Nếu không tìm thấy window handle, thử tìm tất cả
                //MessageBox.Show(
                //    "Ứng dụng đã chạy nhưng không thể kích hoạt!",
                //    "Thông báo",
                //    MessageBoxButtons.OK,
                //    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Lỗi khi kích hoạt ứng dụng: {ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Kiểm tra ứng dụng có đang chạy từ Task Scheduler không
        /// </summary>
        private static bool IsRunningFromTaskScheduler()
        {
            // ✅ Kiểm tra tham số dòng lệnh -autostart
            string[] args = Environment.GetCommandLineArgs();
            foreach (string arg in args)
            {
                if (arg.Equals("-autostart", StringComparison.OrdinalIgnoreCase) ||
                    arg.Equals("/autostart", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            // Kiểm tra user là SYSTEM
            if (Environment.UserName.Equals("SYSTEM", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // Kiểm tra parent process
            try
            {
                using (Process currentProcess = Process.GetCurrentProcess())
                using (Process parentProcess = GetParentProcess(currentProcess.Id))
                {
                    if (parentProcess != null)
                    {
                        string parentName = parentProcess.ProcessName.ToLower();

                        string[] schedulerProcesses = {
                            "svchost", "taskeng", "taskschd",
                            "taskhostw", "taskhostex", "taskhost"
                        };

                        foreach (string name in schedulerProcesses)
                        {
                            if (parentName.Contains(name))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            catch { }

            // Kiểm tra biến môi trường
            string taskTrigger = Environment.GetEnvironmentVariable("TASK_TRIGGER");
            string taskName = Environment.GetEnvironmentVariable("TASK_NAME");

            if (!string.IsNullOrEmpty(taskTrigger) || !string.IsNullOrEmpty(taskName))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Lấy Process cha
        /// </summary>
        private static Process GetParentProcess(int id)
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
    }
}