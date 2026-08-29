using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ToolTaiHD
{
    public partial class LoadingForm : DevExpress.XtraEditors.XtraForm
    {
        public LoadingForm()
        {
            InitializeComponent();
            this.Shown += async (s, e) =>
            {
                // Đảm bảo form đã hiển thị hoàn toàn
                await Task.Delay(50);
            };
        }

        private async void LoadingForm_Load(object sender, EventArgs e)
        {
            this.Refresh();
            Application.DoEvents();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //if (progressBar1.Value < 100)
            //{
            //    // Tăng nhưng không vượt quá 100
            //    progressBar1.Value = Math.Min(progressBar1.Value + 50, 100);
            //    Application.DoEvents();
            //}
            //else
            //{
            //    timer1.Stop(); 
            //}
        }

        private void progressPanel1_Click(object sender, EventArgs e)
        {

        }
    }
}