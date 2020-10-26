namespace EXPORT_SHEET_FROM_EXCEL
{
    partial class Frm_Main
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Frm_Main));
            this.label1 = new System.Windows.Forms.Label();
            this.txt_file_excel_path = new System.Windows.Forms.TextBox();
            this.btn_load_file_excel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.lst_sheet_from_excel_file = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cbo_mode_export = new System.Windows.Forms.ComboBox();
            this.btn_export = new System.Windows.Forms.Button();
            this.OpenFileDLG1 = new System.Windows.Forms.OpenFileDialog();
            this.btn_thoat = new System.Windows.Forms.Button();
            this.FolderBrowserDLG1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.Blue;
            this.label1.Location = new System.Drawing.Point(21, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "FILE EXCEL";
            // 
            // txt_file_excel_path
            // 
            this.txt_file_excel_path.BackColor = System.Drawing.Color.White;
            this.txt_file_excel_path.Cursor = System.Windows.Forms.Cursors.Default;
            this.txt_file_excel_path.Location = new System.Drawing.Point(99, 16);
            this.txt_file_excel_path.Name = "txt_file_excel_path";
            this.txt_file_excel_path.ReadOnly = true;
            this.txt_file_excel_path.Size = new System.Drawing.Size(376, 21);
            this.txt_file_excel_path.TabIndex = 0;
            this.txt_file_excel_path.TabStop = false;
            // 
            // btn_load_file_excel
            // 
            this.btn_load_file_excel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_load_file_excel.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_load_file_excel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.btn_load_file_excel.Location = new System.Drawing.Point(23, 278);
            this.btn_load_file_excel.Name = "btn_load_file_excel";
            this.btn_load_file_excel.Size = new System.Drawing.Size(138, 32);
            this.btn_load_file_excel.TabIndex = 1;
            this.btn_load_file_excel.Text = "&NẠP FILE EXCEL";
            this.btn_load_file_excel.UseVisualStyleBackColor = true;
            this.btn_load_file_excel.Click += new System.EventHandler(this.btn_load_file_excel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.Blue;
            this.label2.Location = new System.Drawing.Point(21, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(299, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "DANH SÁCH NHỮNG SHEET CÓ TRONG FILE EXCEL";
            // 
            // lst_sheet_from_excel_file
            // 
            this.lst_sheet_from_excel_file.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lst_sheet_from_excel_file.FormattingEnabled = true;
            this.lst_sheet_from_excel_file.ItemHeight = 14;
            this.lst_sheet_from_excel_file.Location = new System.Drawing.Point(24, 64);
            this.lst_sheet_from_excel_file.Name = "lst_sheet_from_excel_file";
            this.lst_sheet_from_excel_file.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lst_sheet_from_excel_file.Size = new System.Drawing.Size(451, 144);
            this.lst_sheet_from_excel_file.TabIndex = 2;
            this.lst_sheet_from_excel_file.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lst_sheet_from_excel_file_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.Blue;
            this.label3.Location = new System.Drawing.Point(21, 219);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "KIỂU XUẤT";
            // 
            // cbo_mode_export
            // 
            this.cbo_mode_export.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbo_mode_export.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbo_mode_export.FormattingEnabled = true;
            this.cbo_mode_export.IntegralHeight = false;
            this.cbo_mode_export.Location = new System.Drawing.Point(24, 238);
            this.cbo_mode_export.Margin = new System.Windows.Forms.Padding(0);
            this.cbo_mode_export.Name = "cbo_mode_export";
            this.cbo_mode_export.Size = new System.Drawing.Size(451, 24);
            this.cbo_mode_export.TabIndex = 3;
            // 
            // btn_export
            // 
            this.btn_export.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_export.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_export.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.btn_export.Location = new System.Drawing.Point(167, 278);
            this.btn_export.Name = "btn_export";
            this.btn_export.Size = new System.Drawing.Size(235, 32);
            this.btn_export.TabIndex = 4;
            this.btn_export.Text = "&XUẤT NHỮNG SHEET ĐANG CHỌN";
            this.btn_export.UseVisualStyleBackColor = true;
            this.btn_export.Click += new System.EventHandler(this.btn_export_Click);
            // 
            // btn_thoat
            // 
            this.btn_thoat.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_thoat.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btn_thoat.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_thoat.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.btn_thoat.Location = new System.Drawing.Point(409, 278);
            this.btn_thoat.Name = "btn_thoat";
            this.btn_thoat.Size = new System.Drawing.Size(67, 32);
            this.btn_thoat.TabIndex = 5;
            this.btn_thoat.Text = "&THOÁT";
            this.btn_thoat.UseVisualStyleBackColor = true;
            this.btn_thoat.Click += new System.EventHandler(this.btn_thoat_Click);
            // 
            // Frm_Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_thoat;
            this.ClientSize = new System.Drawing.Size(500, 331);
            this.Controls.Add(this.btn_thoat);
            this.Controls.Add(this.btn_export);
            this.Controls.Add(this.cbo_mode_export);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lst_sheet_from_excel_file);
            this.Controls.Add(this.btn_load_file_excel);
            this.Controls.Add(this.txt_file_excel_path);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.Name = "Frm_Main";
            this.Padding = new System.Windows.Forms.Padding(18, 19, 18, 19);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "EXPORT SHEET FROM EXCEL v1.0";
            this.Load += new System.EventHandler(this.Frm_Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_file_excel_path;
        private System.Windows.Forms.Button btn_load_file_excel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox lst_sheet_from_excel_file;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbo_mode_export;
        private System.Windows.Forms.Button btn_export;
        private System.Windows.Forms.OpenFileDialog OpenFileDLG1;
        private System.Windows.Forms.Button btn_thoat;
        private System.Windows.Forms.FolderBrowserDialog FolderBrowserDLG1;
    }
}

