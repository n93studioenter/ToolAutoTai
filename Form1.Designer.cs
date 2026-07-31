namespace ToolTaiHD
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            DevExpress.XtraEditors.Controls.EditorButtonImageOptions editorButtonImageOptions1 = new DevExpress.XtraEditors.Controls.EditorButtonImageOptions();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject1 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject2 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject3 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject4 = new DevExpress.Utils.SerializableAppearanceObject();
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.companyBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDbpath = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colFolderPath = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colMST = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colSTT = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStatus = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colIsRun = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDauvao = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDaura = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemButtonEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit();
            this.congtyBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.btnRun = new DevExpress.XtraEditors.SimpleButton();
            this.groupControl1 = new DevExpress.XtraEditors.GroupControl();
            this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
            this.txttimeout = new DevExpress.XtraEditors.TextEdit();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.txtBlock3 = new DevExpress.XtraEditors.TextEdit();
            this.chkMoc3 = new DevExpress.XtraEditors.CheckEdit();
            this.txtSoluongtai = new DevExpress.XtraEditors.TextEdit();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.txtBlock2 = new DevExpress.XtraEditors.TextEdit();
            this.chkMoc2 = new DevExpress.XtraEditors.CheckEdit();
            this.txtBlock1 = new DevExpress.XtraEditors.TextEdit();
            this.chkMoc1 = new DevExpress.XtraEditors.CheckEdit();
            this.checkEdit1 = new DevExpress.XtraEditors.CheckEdit();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.groupControl2 = new DevExpress.XtraEditors.GroupControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.label1 = new System.Windows.Forms.Label();
            this.txtsolanlap = new DevExpress.XtraEditors.TextEdit();
            this.txtSovongtai = new DevExpress.XtraEditors.TextEdit();
            this.label2 = new System.Windows.Forms.Label();
            this.gridColumn4 = new DevExpress.XtraGrid.Columns.GridColumn();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.companyBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemButtonEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.congtyBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).BeginInit();
            this.groupControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txttimeout.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBlock3.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkMoc3.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSoluongtai.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBlock2.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkMoc2.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBlock1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkMoc1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl2)).BeginInit();
            this.groupControl2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtsolanlap.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSovongtai.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // gridControl1
            // 
            this.gridControl1.DataSource = this.companyBindingSource;
            this.gridControl1.Location = new System.Drawing.Point(8, 153);
            this.gridControl1.MainView = this.gridView1;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemButtonEdit1});
            this.gridControl1.Size = new System.Drawing.Size(1244, 294);
            this.gridControl1.TabIndex = 0;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            this.gridControl1.Click += new System.EventHandler(this.gridControl1_Click);
            // 
            // companyBindingSource
            // 
            this.companyBindingSource.DataSource = typeof(ToolTaiHD.Form1.Company);
            // 
            // gridView1
            // 
            this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colName,
            this.colDbpath,
            this.colFolderPath,
            this.colMST,
            this.colSTT,
            this.colStatus,
            this.colIsRun,
            this.colDauvao,
            this.colDaura,
            this.gridColumn1,
            this.gridColumn2,
            this.gridColumn3,
            this.gridColumn4});
            this.gridView1.GridControl = this.gridControl1;
            this.gridView1.Name = "gridView1";
            this.gridView1.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.gridView1_CellValueChanged);
            // 
            // colName
            // 
            this.colName.FieldName = "Name";
            this.colName.MinWidth = 25;
            this.colName.Name = "colName";
            this.colName.Visible = true;
            this.colName.VisibleIndex = 1;
            this.colName.Width = 217;
            // 
            // colDbpath
            // 
            this.colDbpath.FieldName = "Dbpath";
            this.colDbpath.MinWidth = 25;
            this.colDbpath.Name = "colDbpath";
            this.colDbpath.Width = 164;
            // 
            // colFolderPath
            // 
            this.colFolderPath.Caption = "Đường dẫn thư mục";
            this.colFolderPath.FieldName = "FolderPath";
            this.colFolderPath.MinWidth = 25;
            this.colFolderPath.Name = "colFolderPath";
            this.colFolderPath.Visible = true;
            this.colFolderPath.VisibleIndex = 2;
            this.colFolderPath.Width = 200;
            // 
            // colMST
            // 
            this.colMST.FieldName = "MST";
            this.colMST.MinWidth = 25;
            this.colMST.Name = "colMST";
            this.colMST.Width = 164;
            // 
            // colSTT
            // 
            this.colSTT.FieldName = "STT";
            this.colSTT.MinWidth = 25;
            this.colSTT.Name = "colSTT";
            this.colSTT.Visible = true;
            this.colSTT.VisibleIndex = 0;
            this.colSTT.Width = 48;
            // 
            // colStatus
            // 
            this.colStatus.Caption = "Trạng thái";
            this.colStatus.FieldName = "Status";
            this.colStatus.MinWidth = 25;
            this.colStatus.Name = "colStatus";
            this.colStatus.Visible = true;
            this.colStatus.VisibleIndex = 3;
            this.colStatus.Width = 238;
            // 
            // colIsRun
            // 
            this.colIsRun.FieldName = "IsRun";
            this.colIsRun.MinWidth = 25;
            this.colIsRun.Name = "colIsRun";
            this.colIsRun.Visible = true;
            this.colIsRun.VisibleIndex = 4;
            this.colIsRun.Width = 83;
            // 
            // colDauvao
            // 
            this.colDauvao.FieldName = "Dauvao";
            this.colDauvao.MinWidth = 25;
            this.colDauvao.Name = "colDauvao";
            this.colDauvao.Width = 125;
            // 
            // colDaura
            // 
            this.colDaura.FieldName = "Daura";
            this.colDaura.MinWidth = 25;
            this.colDaura.Name = "colDaura";
            this.colDaura.Width = 134;
            // 
            // gridColumn1
            // 
            this.gridColumn1.Caption = "Folder";
            this.gridColumn1.FieldName = "Folder";
            this.gridColumn1.MinWidth = 25;
            this.gridColumn1.Name = "gridColumn1";
            this.gridColumn1.Visible = true;
            this.gridColumn1.VisibleIndex = 5;
            this.gridColumn1.Width = 109;
            // 
            // gridColumn2
            // 
            this.gridColumn2.Caption = "RunCount";
            this.gridColumn2.FieldName = "RunCount";
            this.gridColumn2.MinWidth = 25;
            this.gridColumn2.Name = "gridColumn2";
            this.gridColumn2.Visible = true;
            this.gridColumn2.VisibleIndex = 6;
            this.gridColumn2.Width = 81;
            // 
            // gridColumn3
            // 
            this.gridColumn3.Caption = "Ngày tài khoản hết hạn";
            this.gridColumn3.FieldName = "DateAccount";
            this.gridColumn3.MinWidth = 25;
            this.gridColumn3.Name = "gridColumn3";
            this.gridColumn3.Visible = true;
            this.gridColumn3.VisibleIndex = 7;
            this.gridColumn3.Width = 183;
            // 
            // repositoryItemButtonEdit1
            // 
            this.repositoryItemButtonEdit1.AutoHeight = false;
            editorButtonImageOptions1.Image = global::ToolTaiHD.Properties.Resources.cancel_32x32;
            this.repositoryItemButtonEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Glyph, "", -1, true, true, false, editorButtonImageOptions1, new DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), serializableAppearanceObject1, serializableAppearanceObject2, serializableAppearanceObject3, serializableAppearanceObject4, "", null, null, DevExpress.Utils.ToolTipAnchor.Default)});
            this.repositoryItemButtonEdit1.Name = "repositoryItemButtonEdit1";
            // 
            // congtyBindingSource
            // 
            this.congtyBindingSource.DataSource = typeof(ToolTaiHD.Form1.Congty);
            // 
            // btnRun
            // 
            this.btnRun.ImageOptions.Image = global::ToolTaiHD.Properties.Resources.play_32x32;
            this.btnRun.Location = new System.Drawing.Point(1206, 464);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(44, 29);
            this.btnRun.TabIndex = 13;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // groupControl1
            // 
            this.groupControl1.Controls.Add(this.labelControl4);
            this.groupControl1.Controls.Add(this.txttimeout);
            this.groupControl1.Controls.Add(this.labelControl2);
            this.groupControl1.Controls.Add(this.txtBlock3);
            this.groupControl1.Controls.Add(this.chkMoc3);
            this.groupControl1.Controls.Add(this.txtSoluongtai);
            this.groupControl1.Controls.Add(this.labelControl1);
            this.groupControl1.Controls.Add(this.radioButton1);
            this.groupControl1.Controls.Add(this.txtBlock2);
            this.groupControl1.Controls.Add(this.chkMoc2);
            this.groupControl1.Controls.Add(this.txtBlock1);
            this.groupControl1.Controls.Add(this.chkMoc1);
            this.groupControl1.Controls.Add(this.checkEdit1);
            this.groupControl1.Location = new System.Drawing.Point(12, 12);
            this.groupControl1.Name = "groupControl1";
            this.groupControl1.Size = new System.Drawing.Size(1238, 109);
            this.groupControl1.TabIndex = 14;
            this.groupControl1.Text = "Thiết lập";
            // 
            // labelControl4
            // 
            this.labelControl4.Appearance.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.labelControl4.Appearance.Options.UseFont = true;
            this.labelControl4.Location = new System.Drawing.Point(17, 73);
            this.labelControl4.Name = "labelControl4";
            this.labelControl4.Size = new System.Drawing.Size(183, 16);
            this.labelControl4.TabIndex = 15;
            this.labelControl4.Text = "Mốc thời gian tải trong ngày";
            // 
            // txttimeout
            // 
            this.txttimeout.EditValue = "6";
            this.txttimeout.Location = new System.Drawing.Point(769, 31);
            this.txttimeout.Name = "txttimeout";
            this.txttimeout.Size = new System.Drawing.Size(67, 24);
            this.txttimeout.TabIndex = 14;
            this.txttimeout.EditValueChanged += new System.EventHandler(this.txttimeout_EditValueChanged);
            // 
            // labelControl2
            // 
            this.labelControl2.Appearance.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.labelControl2.Appearance.ForeColor = System.Drawing.Color.Yellow;
            this.labelControl2.Appearance.Options.UseFont = true;
            this.labelControl2.Appearance.Options.UseForeColor = true;
            this.labelControl2.Location = new System.Drawing.Point(533, 35);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(192, 16);
            this.labelControl2.TabIndex = 13;
            this.labelControl2.Text = "Thời gian chờ tải 1 hoá đơn(s)";
            // 
            // txtBlock3
            // 
            this.txtBlock3.Enabled = false;
            this.txtBlock3.Location = new System.Drawing.Point(769, 69);
            this.txtBlock3.Name = "txtBlock3";
            this.txtBlock3.Size = new System.Drawing.Size(125, 24);
            this.txtBlock3.TabIndex = 12;
            this.txtBlock3.EditValueChanged += new System.EventHandler(this.txtBlock3_EditValueChanged);
            this.txtBlock3.Validated += new System.EventHandler(this.txtBlock3_Validated);
            // 
            // chkMoc3
            // 
            this.chkMoc3.Location = new System.Drawing.Point(679, 73);
            this.chkMoc3.Name = "chkMoc3";
            this.chkMoc3.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.chkMoc3.Properties.Appearance.Options.UseFont = true;
            this.chkMoc3.Properties.Caption = "Mốc TG3";
            this.chkMoc3.Size = new System.Drawing.Size(94, 20);
            this.chkMoc3.TabIndex = 11;
            this.chkMoc3.CheckedChanged += new System.EventHandler(this.chkMoc3_CheckedChanged);
            // 
            // txtSoluongtai
            // 
            this.txtSoluongtai.EditValue = "2";
            this.txtSoluongtai.Location = new System.Drawing.Point(435, 31);
            this.txtSoluongtai.Name = "txtSoluongtai";
            this.txtSoluongtai.Size = new System.Drawing.Size(53, 24);
            this.txtSoluongtai.TabIndex = 10;
            this.txtSoluongtai.EditValueChanged += new System.EventHandler(this.textEdit4_EditValueChanged);
            // 
            // labelControl1
            // 
            this.labelControl1.Appearance.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.labelControl1.Appearance.ForeColor = System.Drawing.Color.Lime;
            this.labelControl1.Appearance.Options.UseFont = true;
            this.labelControl1.Appearance.Options.UseForeColor = true;
            this.labelControl1.Location = new System.Drawing.Point(231, 35);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(187, 16);
            this.labelControl1.TabIndex = 9;
            this.labelControl1.Text = "Số lượng công ty tải cùng lúc";
            this.labelControl1.Click += new System.EventHandler(this.labelControl1_Click);
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(932, 0);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(100, 20);
            this.radioButton1.TabIndex = 7;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "radioButton1";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // txtBlock2
            // 
            this.txtBlock2.Enabled = false;
            this.txtBlock2.Location = new System.Drawing.Point(548, 71);
            this.txtBlock2.Name = "txtBlock2";
            this.txtBlock2.Size = new System.Drawing.Size(125, 24);
            this.txtBlock2.TabIndex = 4;
            this.txtBlock2.EditValueChanged += new System.EventHandler(this.txtBlock2_EditValueChanged);
            this.txtBlock2.Validated += new System.EventHandler(this.txtBlock2_Validated);
            // 
            // chkMoc2
            // 
            this.chkMoc2.Location = new System.Drawing.Point(454, 73);
            this.chkMoc2.Name = "chkMoc2";
            this.chkMoc2.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.chkMoc2.Properties.Appearance.Options.UseFont = true;
            this.chkMoc2.Properties.Caption = "Mốc TG2";
            this.chkMoc2.Size = new System.Drawing.Size(94, 20);
            this.chkMoc2.TabIndex = 3;
            this.chkMoc2.CheckedChanged += new System.EventHandler(this.chkMoc2_CheckedChanged);
            // 
            // txtBlock1
            // 
            this.txtBlock1.Enabled = false;
            this.txtBlock1.Location = new System.Drawing.Point(306, 69);
            this.txtBlock1.Name = "txtBlock1";
            this.txtBlock1.Size = new System.Drawing.Size(125, 24);
            this.txtBlock1.TabIndex = 2;
            this.txtBlock1.EditValueChanged += new System.EventHandler(this.txtBlock1_EditValueChanged);
            this.txtBlock1.Validated += new System.EventHandler(this.txtBlock1_Validated);
            // 
            // chkMoc1
            // 
            this.chkMoc1.Location = new System.Drawing.Point(223, 73);
            this.chkMoc1.Name = "chkMoc1";
            this.chkMoc1.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.chkMoc1.Properties.Appearance.Options.UseFont = true;
            this.chkMoc1.Properties.Caption = "Mốc TG1";
            this.chkMoc1.Size = new System.Drawing.Size(94, 20);
            this.chkMoc1.TabIndex = 1;
            this.chkMoc1.CheckedChanged += new System.EventHandler(this.chkMoc1_CheckedChanged);
            // 
            // checkEdit1
            // 
            this.checkEdit1.Location = new System.Drawing.Point(17, 33);
            this.checkEdit1.Name = "checkEdit1";
            this.checkEdit1.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.checkEdit1.Properties.Appearance.ForeColor = System.Drawing.Color.LightCoral;
            this.checkEdit1.Properties.Appearance.Options.UseFont = true;
            this.checkEdit1.Properties.Appearance.Options.UseForeColor = true;
            this.checkEdit1.Properties.Caption = "Chạy khi khởi động máy";
            this.checkEdit1.Size = new System.Drawing.Size(208, 20);
            this.checkEdit1.TabIndex = 0;
            this.checkEdit1.CheckedChanged += new System.EventHandler(this.checkEdit1_CheckedChanged);
            // 
            // richTextBox1
            // 
            this.richTextBox1.BackColor = System.Drawing.Color.SlateGray;
            this.richTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox1.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.richTextBox1.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.richTextBox1.Location = new System.Drawing.Point(2, 27);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(1240, 98);
            this.richTextBox1.TabIndex = 15;
            this.richTextBox1.Text = "...";
            // 
            // groupControl2
            // 
            this.groupControl2.Controls.Add(this.richTextBox1);
            this.groupControl2.Location = new System.Drawing.Point(8, 510);
            this.groupControl2.Name = "groupControl2";
            this.groupControl2.Size = new System.Drawing.Size(1244, 127);
            this.groupControl2.TabIndex = 16;
            this.groupControl2.Text = "Ghi chú Log";
            // 
            // labelControl3
            // 
            this.labelControl3.Location = new System.Drawing.Point(13, 127);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(12, 16);
            this.labelControl3.TabIndex = 17;
            this.labelControl3.Text = "...";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label1.Location = new System.Drawing.Point(869, 131);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(175, 16);
            this.label1.TabIndex = 18;
            this.label1.Text = "Số lần lặp lại trên 1 lần tải";
            // 
            // txtsolanlap
            // 
            this.txtsolanlap.EditValue = "2";
            this.txtsolanlap.Location = new System.Drawing.Point(1058, 127);
            this.txtsolanlap.Name = "txtsolanlap";
            this.txtsolanlap.Size = new System.Drawing.Size(125, 24);
            this.txtsolanlap.TabIndex = 19;
            // 
            // txtSovongtai
            // 
            this.txtSovongtai.EditValue = "2";
            this.txtSovongtai.Location = new System.Drawing.Point(719, 127);
            this.txtSovongtai.Name = "txtSovongtai";
            this.txtSovongtai.Size = new System.Drawing.Size(125, 24);
            this.txtSovongtai.TabIndex = 21;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(590, 131);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(123, 16);
            this.label2.TabIndex = 20;
            this.label2.Text = "Số vòng tải tất cả";
            // 
            // gridColumn4
            // 
            this.gridColumn4.Caption = "Clear";
            this.gridColumn4.FieldName = "Clear";
            this.gridColumn4.MinWidth = 25;
            this.gridColumn4.Name = "gridColumn4";
            this.gridColumn4.Visible = true;
            this.gridColumn4.VisibleIndex = 8;
            this.gridColumn4.Width = 63;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1264, 649);
            this.Controls.Add(this.txtSovongtai);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtsolanlap);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.labelControl3);
            this.Controls.Add(this.groupControl2);
            this.Controls.Add(this.groupControl1);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.gridControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Vietstar Auto Download";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.companyBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemButtonEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.congtyBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).EndInit();
            this.groupControl1.ResumeLayout(false);
            this.groupControl1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txttimeout.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBlock3.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkMoc3.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSoluongtai.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBlock2.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkMoc2.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBlock1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkMoc1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl2)).EndInit();
            this.groupControl2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txtsolanlap.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSovongtai.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private System.Windows.Forms.BindingSource congtyBindingSource;
        private DevExpress.XtraEditors.SimpleButton btnRun;
        private DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit repositoryItemButtonEdit1;
        private DevExpress.XtraEditors.GroupControl groupControl1;
        private DevExpress.XtraEditors.TextEdit txtBlock2;
        private DevExpress.XtraEditors.CheckEdit chkMoc2;
        private DevExpress.XtraEditors.TextEdit txtBlock1;
        private DevExpress.XtraEditors.CheckEdit chkMoc1;
        private DevExpress.XtraEditors.CheckEdit checkEdit1;
        private System.Windows.Forms.BindingSource companyBindingSource;
        private DevExpress.XtraGrid.Columns.GridColumn colName;
        private DevExpress.XtraGrid.Columns.GridColumn colDbpath;
        private DevExpress.XtraGrid.Columns.GridColumn colFolderPath;
        private DevExpress.XtraGrid.Columns.GridColumn colMST;
        private DevExpress.XtraGrid.Columns.GridColumn colSTT;
        private DevExpress.XtraGrid.Columns.GridColumn colStatus;
        private DevExpress.XtraGrid.Columns.GridColumn colIsRun;
        private DevExpress.XtraGrid.Columns.GridColumn colDauvao;
        private DevExpress.XtraGrid.Columns.GridColumn colDaura;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private DevExpress.XtraEditors.TextEdit txtSoluongtai;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.TextEdit txtBlock3;
        private DevExpress.XtraEditors.CheckEdit chkMoc3;
        private DevExpress.XtraEditors.GroupControl groupControl2;
        private DevExpress.XtraEditors.TextEdit txttimeout;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
        private System.Windows.Forms.Label label1;
        private DevExpress.XtraEditors.TextEdit txtsolanlap;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        private DevExpress.XtraEditors.TextEdit txtSovongtai;
        private System.Windows.Forms.Label label2;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn4;
    }
}

