using DevExpress.LookAndFeel;
using DevExpress.Skins;
using DevExpress.UserSkins;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Windows.Forms;

namespace ToolTaiHD
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            ServicePointManager.ServerCertificateValidationCallback = (a, b, c, d) => true;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // ✅ Kiểm tra AutoStart
            bool isAutoStart = args.Any(a => a.Equals("-autostart", StringComparison.OrdinalIgnoreCase));

            // ✅ Kiểm tra Task Scheduler
            bool isTaskScheduler = IsRunningFromTaskScheduler();
           // isAutoStart = true;
            // ✅ Tạo form với tham số isAutoMode
            Form1 form = new Form1(isAutoStart || isTaskScheduler);

            Application.Run(form);
        }

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