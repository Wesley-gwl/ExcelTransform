namespace ExcelTransform
{
    partial class FrmMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.OFDExcel = new System.Windows.Forms.OpenFileDialog();
            this.DGVExport = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.BtnSelectSKU = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.TxtSKUUrl = new System.Windows.Forms.TextBox();
            this.TxtBuyUrl = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnOpen = new System.Windows.Forms.Button();
            this.tbtUrl = new System.Windows.Forms.TextBox();
            this.BtnTransfrom = new System.Windows.Forms.Button();
            this.BtnExportSuccess = new System.Windows.Forms.Button();
            this.tab1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.LblTotalLine = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.BtnSplitSuccessExport = new System.Windows.Forms.Button();
            this.LblSuccessLine = new System.Windows.Forms.Label();
            this.DGVExportSuccess = new System.Windows.Forms.DataGridView();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.panel3 = new System.Windows.Forms.Panel();
            this.LblFailLine = new System.Windows.Forms.Label();
            this.BtnExportFail = new System.Windows.Forms.Button();
            this.DGVExportFail = new System.Windows.Forms.DataGridView();
            this.OFDSKUtxt = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.DGVExport)).BeginInit();
            this.panel1.SuspendLayout();
            this.tab1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGVExportSuccess)).BeginInit();
            this.tabPage3.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGVExportFail)).BeginInit();
            this.SuspendLayout();
            // 
            // OFDExcel
            // 
            this.OFDExcel.FileName = "OFDExcel";
            // 
            // DGVExport
            // 
            this.DGVExport.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVExport.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.DGVExport.Location = new System.Drawing.Point(2, 37);
            this.DGVExport.Margin = new System.Windows.Forms.Padding(2);
            this.DGVExport.Name = "DGVExport";
            this.DGVExport.RowTemplate.Height = 27;
            this.DGVExport.Size = new System.Drawing.Size(738, 284);
            this.DGVExport.TabIndex = 3;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.BtnSelectSKU);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.TxtSKUUrl);
            this.panel1.Controls.Add(this.TxtBuyUrl);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.BtnOpen);
            this.panel1.Controls.Add(this.tbtUrl);
            this.panel1.Controls.Add(this.BtnTransfrom);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(750, 156);
            this.panel1.TabIndex = 4;
            // 
            // BtnSelectSKU
            // 
            this.BtnSelectSKU.Location = new System.Drawing.Point(380, 90);
            this.BtnSelectSKU.Margin = new System.Windows.Forms.Padding(2);
            this.BtnSelectSKU.Name = "BtnSelectSKU";
            this.BtnSelectSKU.Size = new System.Drawing.Size(106, 22);
            this.BtnSelectSKU.TabIndex = 12;
            this.BtnSelectSKU.Text = "选择SKU规则文件";
            this.BtnSelectSKU.UseVisualStyleBackColor = true;
            this.BtnSelectSKU.Click += new System.EventHandler(this.BtnSelectSKU_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(33, 93);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 12);
            this.label3.TabIndex = 11;
            this.label3.Text = "SKU规则文件：";
            // 
            // TxtSKUUrl
            // 
            this.TxtSKUUrl.Location = new System.Drawing.Point(117, 90);
            this.TxtSKUUrl.Margin = new System.Windows.Forms.Padding(2);
            this.TxtSKUUrl.Name = "TxtSKUUrl";
            this.TxtSKUUrl.ReadOnly = true;
            this.TxtSKUUrl.Size = new System.Drawing.Size(216, 21);
            this.TxtSKUUrl.TabIndex = 10;
            // 
            // TxtBuyUrl
            // 
            this.TxtBuyUrl.Location = new System.Drawing.Point(117, 54);
            this.TxtBuyUrl.Margin = new System.Windows.Forms.Padding(2);
            this.TxtBuyUrl.Name = "TxtBuyUrl";
            this.TxtBuyUrl.Size = new System.Drawing.Size(216, 21);
            this.TxtBuyUrl.TabIndex = 9;
            this.TxtBuyUrl.Text = "https://item.taobao.com/item.htm?id=527446715774&ali_trackid=2:mm_14507504_340547" +
    "7_35092123:1583931338_108_1158122964&spm=1002.8113010.1999451587.1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(51, 56);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "商品链接：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(51, 26);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "文件路径：";
            // 
            // BtnOpen
            // 
            this.BtnOpen.Location = new System.Drawing.Point(380, 21);
            this.BtnOpen.Margin = new System.Windows.Forms.Padding(2);
            this.BtnOpen.Name = "BtnOpen";
            this.BtnOpen.Size = new System.Drawing.Size(106, 22);
            this.BtnOpen.TabIndex = 5;
            this.BtnOpen.Text = "选择导入文件";
            this.BtnOpen.UseVisualStyleBackColor = true;
            this.BtnOpen.Click += new System.EventHandler(this.BtnOpen_Click);
            // 
            // tbtUrl
            // 
            this.tbtUrl.Location = new System.Drawing.Point(117, 21);
            this.tbtUrl.Margin = new System.Windows.Forms.Padding(2);
            this.tbtUrl.Name = "tbtUrl";
            this.tbtUrl.ReadOnly = true;
            this.tbtUrl.Size = new System.Drawing.Size(216, 21);
            this.tbtUrl.TabIndex = 4;
            // 
            // BtnTransfrom
            // 
            this.BtnTransfrom.Location = new System.Drawing.Point(519, 21);
            this.BtnTransfrom.Margin = new System.Windows.Forms.Padding(2);
            this.BtnTransfrom.Name = "BtnTransfrom";
            this.BtnTransfrom.Size = new System.Drawing.Size(106, 22);
            this.BtnTransfrom.TabIndex = 3;
            this.BtnTransfrom.Text = "转换";
            this.BtnTransfrom.UseVisualStyleBackColor = true;
            this.BtnTransfrom.Click += new System.EventHandler(this.BtnTransfrom_Click);
            // 
            // BtnExportSuccess
            // 
            this.BtnExportSuccess.Location = new System.Drawing.Point(4, 2);
            this.BtnExportSuccess.Margin = new System.Windows.Forms.Padding(2);
            this.BtnExportSuccess.Name = "BtnExportSuccess";
            this.BtnExportSuccess.Size = new System.Drawing.Size(106, 31);
            this.BtnExportSuccess.TabIndex = 6;
            this.BtnExportSuccess.Text = "导出";
            this.BtnExportSuccess.UseVisualStyleBackColor = true;
            this.BtnExportSuccess.Click += new System.EventHandler(this.BtnExportSuccess_Click);
            // 
            // tab1
            // 
            this.tab1.Controls.Add(this.tabPage1);
            this.tab1.Controls.Add(this.tabPage2);
            this.tab1.Controls.Add(this.tabPage3);
            this.tab1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tab1.Location = new System.Drawing.Point(0, 156);
            this.tab1.Margin = new System.Windows.Forms.Padding(2);
            this.tab1.Name = "tab1";
            this.tab1.SelectedIndex = 0;
            this.tab1.Size = new System.Drawing.Size(750, 349);
            this.tab1.TabIndex = 5;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.LblTotalLine);
            this.tabPage1.Controls.Add(this.DGVExport);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage1.Size = new System.Drawing.Size(742, 323);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "导入数据";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // LblTotalLine
            // 
            this.LblTotalLine.AutoSize = true;
            this.LblTotalLine.Location = new System.Drawing.Point(19, 15);
            this.LblTotalLine.Name = "LblTotalLine";
            this.LblTotalLine.Size = new System.Drawing.Size(71, 12);
            this.LblTotalLine.TabIndex = 13;
            this.LblTotalLine.Text = "当前总共0行";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.panel2);
            this.tabPage2.Controls.Add(this.DGVExportSuccess);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage2.Size = new System.Drawing.Size(742, 323);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "转换成功数据";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.BtnSplitSuccessExport);
            this.panel2.Controls.Add(this.LblSuccessLine);
            this.panel2.Controls.Add(this.BtnExportSuccess);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(2, 2);
            this.panel2.Margin = new System.Windows.Forms.Padding(2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(738, 36);
            this.panel2.TabIndex = 1;
            // 
            // BtnSplitSuccessExport
            // 
            this.BtnSplitSuccessExport.Location = new System.Drawing.Point(114, 3);
            this.BtnSplitSuccessExport.Margin = new System.Windows.Forms.Padding(2);
            this.BtnSplitSuccessExport.Name = "BtnSplitSuccessExport";
            this.BtnSplitSuccessExport.Size = new System.Drawing.Size(106, 31);
            this.BtnSplitSuccessExport.TabIndex = 8;
            this.BtnSplitSuccessExport.Text = "分批导出";
            this.BtnSplitSuccessExport.UseVisualStyleBackColor = true;
            this.BtnSplitSuccessExport.Click += new System.EventHandler(this.BtnSplitSuccessExport_Click);
            // 
            // LblSuccessLine
            // 
            this.LblSuccessLine.AutoSize = true;
            this.LblSuccessLine.Location = new System.Drawing.Point(225, 11);
            this.LblSuccessLine.Name = "LblSuccessLine";
            this.LblSuccessLine.Size = new System.Drawing.Size(71, 12);
            this.LblSuccessLine.TabIndex = 7;
            this.LblSuccessLine.Text = "当前总共0行";
            // 
            // DGVExportSuccess
            // 
            this.DGVExportSuccess.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVExportSuccess.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.DGVExportSuccess.Location = new System.Drawing.Point(2, 35);
            this.DGVExportSuccess.Margin = new System.Windows.Forms.Padding(2);
            this.DGVExportSuccess.Name = "DGVExportSuccess";
            this.DGVExportSuccess.RowTemplate.Height = 27;
            this.DGVExportSuccess.Size = new System.Drawing.Size(738, 286);
            this.DGVExportSuccess.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.panel3);
            this.tabPage3.Controls.Add(this.DGVExportFail);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage3.Size = new System.Drawing.Size(742, 323);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "转换失败数据";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.LblFailLine);
            this.panel3.Controls.Add(this.BtnExportFail);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(2, 2);
            this.panel3.Margin = new System.Windows.Forms.Padding(2);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(738, 38);
            this.panel3.TabIndex = 1;
            // 
            // LblFailLine
            // 
            this.LblFailLine.AutoSize = true;
            this.LblFailLine.Location = new System.Drawing.Point(128, 11);
            this.LblFailLine.Name = "LblFailLine";
            this.LblFailLine.Size = new System.Drawing.Size(71, 12);
            this.LblFailLine.TabIndex = 12;
            this.LblFailLine.Text = "当前总共0行";
            // 
            // BtnExportFail
            // 
            this.BtnExportFail.Location = new System.Drawing.Point(4, 2);
            this.BtnExportFail.Margin = new System.Windows.Forms.Padding(2);
            this.BtnExportFail.Name = "BtnExportFail";
            this.BtnExportFail.Size = new System.Drawing.Size(106, 31);
            this.BtnExportFail.TabIndex = 11;
            this.BtnExportFail.Text = "导出";
            this.BtnExportFail.UseVisualStyleBackColor = true;
            this.BtnExportFail.Click += new System.EventHandler(this.BtnExportFail_Click);
            // 
            // DGVExportFail
            // 
            this.DGVExportFail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGVExportFail.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.DGVExportFail.Location = new System.Drawing.Point(2, 36);
            this.DGVExportFail.Margin = new System.Windows.Forms.Padding(2);
            this.DGVExportFail.Name = "DGVExportFail";
            this.DGVExportFail.RowTemplate.Height = 27;
            this.DGVExportFail.Size = new System.Drawing.Size(738, 285);
            this.DGVExportFail.TabIndex = 0;
            // 
            // OFDSKUtxt
            // 
            this.OFDSKUtxt.FileName = "OFDSKUtxt";
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(750, 505);
            this.Controls.Add(this.tab1);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "转换器";
            ((System.ComponentModel.ISupportInitialize)(this.DGVExport)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGVExportSuccess)).EndInit();
            this.tabPage3.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGVExportFail)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog OFDExcel;
        private System.Windows.Forms.DataGridView DGVExport;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnExportSuccess;
        private System.Windows.Forms.Button BtnOpen;
        private System.Windows.Forms.TextBox tbtUrl;
        private System.Windows.Forms.Button BtnTransfrom;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TxtBuyUrl;
        private System.Windows.Forms.TabControl tab1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.DataGridView DGVExportSuccess;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataGridView DGVExportFail;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button BtnExportFail;
        private System.Windows.Forms.Label LblSuccessLine;
        private System.Windows.Forms.Label LblFailLine;
        private System.Windows.Forms.Label LblTotalLine;
        private System.Windows.Forms.Button BtnSplitSuccessExport;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox TxtSKUUrl;
        private System.Windows.Forms.Button BtnSelectSKU;
        private System.Windows.Forms.OpenFileDialog OFDSKUtxt;
    }
}

