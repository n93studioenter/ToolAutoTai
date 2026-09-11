 
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
            
        }

        private async void LoadingForm_Load(object sender, EventArgs e)
        {
          
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            // Tăng 20% mỗi giây
            //progressBarControl1.Position += 20;

            //// Khi đạt hoặc vượt 100% thì dừng lại
            //if (progressBarControl1.Position >= progressBarControl1.Properties.Maximum)
            //{
            //    progressBarControl1.Position = progressBarControl1.Properties.Maximum; // đảm bảo không vượt
            //    timer1.Stop();
                 
            //}
        }

        private void progressPanel1_Click(object sender, EventArgs e)
        {

        }
    }
}