using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;

namespace EXPORT_SHEET_FROM_EXCEL
{
    public partial class Frm_Main : Form
    {
        string File_Excel_Name = "";

        public Frm_Main() { InitializeComponent(); }

        private void Make_Ds_Mode_Export()
        {
            DataTable DTable_Temp = new DataTable();
            DTable_Temp.Columns.Add("ID", typeof(int));
            DTable_Temp.Columns.Add("NAME", typeof(string));

            DTable_Temp.Rows.Add(1, "XUẤT CHUNG MỘT TẬP TIN EXCEL");
            DTable_Temp.Rows.Add(2, "XUẤT RIÊNG NHIỀU TẬP TIN EXCEL");

            cbo_mode_export.DataSource = DTable_Temp;
            cbo_mode_export.DisplayMember = "NAME";
            cbo_mode_export.ValueMember = "ID";
        }

        private void Frm_Main_Load(object sender, EventArgs e) { Make_Ds_Mode_Export(); }

        private void btn_load_file_excel_Click(object sender, EventArgs e)
        {
            OpenFileDLG1.FileName = "";
            OpenFileDLG1.Filter = "File Excel|*.xls;*.xlsx";

            OpenFileDLG1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            OpenFileDLG1.RestoreDirectory = true;

            if (OpenFileDLG1.ShowDialog() == System.Windows.Forms.DialogResult.Cancel) { return; }

            string File_Excel_Path = OpenFileDLG1.FileName.Trim();

            if (File_Excel_Path == "") { return; }

            if (!System.IO.File.Exists(File_Excel_Path)) { MessageBox.Show("TẬP TIN BẠN CHỌN KHÔNG TỒN TẠI", "THÔNG BÁO"); return; }

            Form.ActiveForm.Refresh();
            lst_sheet_from_excel_file.Items.Clear();

            txt_file_excel_path.Text = File_Excel_Path;
            File_Excel_Name = OpenFileDLG1.SafeFileName.Trim();

            // CHECK FILE EXT SUPPORT

            if (Get_File_Ext_From_String(File_Excel_Name).ToLower() != ".xls" && Get_File_Ext_From_String(File_Excel_Name).ToLower() != ".xlsx")
            {
                MessageBox.Show("TẬP TIN BẠN CHỌN KHÔNG ĐƯỢC HỔ TRỢ", "THÔNG BÁO"); return;
            }

            // XỬ LÝ LẤY DANH SÁCH SHEET CỦA FILE EXCEL ĐÃ CHỌN

            try
            {
                FileStream File_Excel = new FileStream(File_Excel_Path, FileMode.Open, FileAccess.Read);

                // NẾU FILE EXCEL PHIÊN BẢN CŨ

                if (Get_File_Ext_From_String(File_Excel_Name).ToLower() == ".xls")
                {
                    HSSFWorkbook WorkBook = new HSSFWorkbook(File_Excel);
                    for (int i = 0; i < WorkBook.NumberOfSheets; i++)
                    {
                        if (WorkBook.GetSheetName(i) != "") { lst_sheet_from_excel_file.Items.Add(WorkBook.GetSheetName(i)); }
                    }
                }

                // NẾU FILE EXCEL PHIÊN BẢN MỚI

                if (Get_File_Ext_From_String(File_Excel_Name).ToLower() == ".xlsx")
                {
                    ExcelPackage ExcelPck = new ExcelPackage(File_Excel);
                    for (int i = 1; i <= ExcelPck.Workbook.Worksheets.Count; i++)
                    {
                        if (ExcelPck.Workbook.Worksheets[i].Name != "") { lst_sheet_from_excel_file.Items.Add(ExcelPck.Workbook.Worksheets[i].Name); }
                    }
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.ToString());
                if (Ex.Message.ToLower().IndexOf("the process cannot access the file") != -1 && Ex.Message.ToLower().IndexOf("because it is being used by another process") != -1)
                {
                    MessageBox.Show("KHÔNG THỂ MỞ DỮ LIỆU\r\n\r\nTẬP TIN BẠN CHỌN ĐANG ĐƯỢC SỬ DỤNG BỞI CHƯƠNG TRÌNH KHÁC", "THÔNG BÁO");
                    return;
                }
                MessageBox.Show(Ex.ToString(), "THÔNG BÁO");
                return;
            }

            if (lst_sheet_from_excel_file.Items.Count == 0) { MessageBox.Show("FILE EXCEL BẠN CHỌN KHÔNG CÓ SHEET NÀO", "THÔNG BÁO"); return; }
        }

        private void btn_thoat_Click(object sender, EventArgs e) { this.Close(); }

        private void btn_export_Click(object sender, EventArgs e)
        {
            if (lst_sheet_from_excel_file.SelectedItems.Count == 0) { MessageBox.Show("BẠN CHƯA CHỌN NHỮNG SHEET CẦN XUẤT", "THÔNG BÁO"); return; }

            // XỬ LÝ THƯ MỤC LƯU TRỮ
            
            FolderBrowserDLG1.RootFolder = Environment.SpecialFolder.Desktop;
            FolderBrowserDLG1.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (FolderBrowserDLG1.ShowDialog() == DialogResult.Cancel) { return; }

            string Folder_Save = FolderBrowserDLG1.SelectedPath.Trim();

            if (Folder_Save == "") { return; }

            if (System.IO.Directory.Exists(Folder_Save) == false) { MessageBox.Show("THƯ MỤC BẠN CHỌN ĐỂ LƯU TRỮ KHÔNG TỒN TẠI", "THÔNG BÁO"); return; }

            Folder_Save = Folder_Save + "\\";
            Folder_Save = Folder_Save.Replace("\\\\", "\\");

            // CHECK FILE EXT SUPPORT

            if (Get_File_Ext_From_String(File_Excel_Name).ToLower() != ".xls" && Get_File_Ext_From_String(File_Excel_Name).ToLower() != ".xlsx")
            {
                MessageBox.Show("TẬP TIN BẠN CHỌN KHÔNG ĐƯỢC HỔ TRỢ", "THÔNG BÁO"); return;
            }

            // TIẾN HÀNH XUẤT SHEET

            try
            {
                FileStream File_Excel = new FileStream(txt_file_excel_path.Text.Trim(), FileMode.Open, FileAccess.Read);

                // NẾU FILE EXCEL PHIÊN BẢN CŨ

                if (Get_File_Ext_From_String(File_Excel_Name).ToLower() == ".xls")
                {
                    HSSFWorkbook WorkBook_Source = new HSSFWorkbook(File_Excel);

                    if (cbo_mode_export.SelectedValue.ToString() == "1")
                    {
                        HSSFWorkbook WorkBook_New = new HSSFWorkbook();

                        // TẠO TÊN FILE EXCEL

                        string File_Excel_Name_New = Random_From_Now();

                        if (File_Excel_Name.IndexOf(".") > -1) { File_Excel_Name_New = File_Excel_Name.Replace(".", " - " + Random_From_Now() + "."); }
                        else { File_Excel_Name_New = File_Excel_Name + " - " + Random_From_Now() + ".xls"; }

                        File_Excel_Name_New = Process_Name_File(File_Excel_Name_New);

                        // TIẾN HÀNH COPY SHEET

                        foreach (var i in lst_sheet_from_excel_file.SelectedItems)
                        {
                            HSSFSheet SHEET_Source = WorkBook_Source.GetSheet(i.ToString()) as HSSFSheet;
                            SHEET_Source.CopyTo(WorkBook_New, SHEET_Source.SheetName, true, true);
                        }

                        // TIẾN HÀNH GHI FILE

                        System.IO.FileStream XFile = new System.IO.FileStream(Folder_Save + File_Excel_Name_New, System.IO.FileMode.Create, System.IO.FileAccess.Write);
                        WorkBook_New.Write(XFile);
                        XFile.Close();
                    }

                    if (cbo_mode_export.SelectedValue.ToString() == "2")
                    {
                        foreach (var i in lst_sheet_from_excel_file.SelectedItems)
                        {
                            HSSFWorkbook WorkBook_New = new HSSFWorkbook();

                            // TẠO TÊN FILE EXCEL

                            string File_Excel_Name_New = i.ToString() + " - " + Random_From_Now() + ".xls";
                            File_Excel_Name_New = Process_Name_File(File_Excel_Name_New);

                            // TIẾN HÀNH COPY SHEET

                            HSSFSheet SHEET_Source = WorkBook_Source.GetSheet(i.ToString()) as HSSFSheet;
                            SHEET_Source.CopyTo(WorkBook_New, SHEET_Source.SheetName, true, true);

                            // TIẾN HÀNH GHI FILE

                            System.IO.FileStream XFile = new System.IO.FileStream(Folder_Save + File_Excel_Name_New, System.IO.FileMode.Create, System.IO.FileAccess.Write);
                            WorkBook_New.Write(XFile);
                            XFile.Close();
                        }
                    }
                }

                // NẾU FILE EXCEL PHIÊN BẢN MỚI

                if (Get_File_Ext_From_String(File_Excel_Name).ToLower() == ".xlsx")
                {
                    ExcelPackage Excel_Pack_Source = new ExcelPackage(File_Excel);

                    if (cbo_mode_export.SelectedValue.ToString() == "1")
                    {
                        // TẠO TÊN FILE EXCEL

                        string File_Excel_Name_New = Random_From_Now();

                        if (File_Excel_Name.IndexOf(".") > -1) { File_Excel_Name_New = File_Excel_Name.Replace(".", " - " + Random_From_Now() + "."); }
                        else { File_Excel_Name_New = File_Excel_Name + " - " + Random_From_Now() + ".xlsx"; }

                        File_Excel_Name_New = Process_Name_File(File_Excel_Name_New);

                        // TẠO FILE EXCEL MỚI

                        System.IO.FileStream XFile = new System.IO.FileStream(Folder_Save + File_Excel_Name_New, System.IO.FileMode.Create, System.IO.FileAccess.Write);
                        ExcelPackage Excel_Pack_New = new ExcelPackage(XFile);

                        // TIẾN HÀNH COPY SHEET

                        foreach (var i in lst_sheet_from_excel_file.SelectedItems)
                        {
                            string Name_Sheet = i.ToString();
                            ExcelWorksheet Excel_WS_Source = Excel_Pack_Source.Workbook.Worksheets[Name_Sheet];
                            if (Name_Sheet.Trim() == "") { Name_Sheet = "_"; }
                            ExcelWorksheet Excel_WS_New = Excel_Pack_New.Workbook.Worksheets.Add(Name_Sheet, Excel_WS_Source);
                        }

                        Excel_Pack_New.Save();

                        XFile.Close();
                    }

                    if (cbo_mode_export.SelectedValue.ToString() == "2")
                    {
                        foreach (var i in lst_sheet_from_excel_file.SelectedItems)
                        {
                            string Name_Sheet = i.ToString();

                            // TẠO TÊN FILE EXCEL

                            string File_Excel_Name_New = Name_Sheet + " - " + Random_From_Now() + ".xlsx";
                            File_Excel_Name_New = Process_Name_File(File_Excel_Name_New);

                            // TẠO FILE EXCEL MỚI

                            System.IO.FileStream XFile = new System.IO.FileStream(Folder_Save + File_Excel_Name_New, System.IO.FileMode.Create, System.IO.FileAccess.Write);
                            ExcelPackage Excel_Pack_New = new ExcelPackage(XFile);

                            // TIẾN HÀNH COPY SHEET

                            ExcelWorksheet Excel_WS_Source = Excel_Pack_Source.Workbook.Worksheets[Name_Sheet];
                            if (Name_Sheet.Trim() == "") { Name_Sheet = "_"; }
                            ExcelWorksheet Excel_WS_New = Excel_Pack_New.Workbook.Worksheets.Add(Name_Sheet, Excel_WS_Source);

                            Excel_Pack_New.Save();

                            XFile.Close();
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.ToString());
                if (Ex.Message.ToLower().IndexOf("the process cannot access the file") != -1 && Ex.Message.ToLower().IndexOf("because it is being used by another process") != -1)
                {
                    MessageBox.Show("KHÔNG THỂ MỞ DỮ LIỆU\r\n\r\nTẬP TIN BẠN CHỌN ĐANG ĐƯỢC SỬ DỤNG BỞI CHƯƠNG TRÌNH KHÁC", "THÔNG BÁO");
                    return;
                }
                MessageBox.Show(Ex.ToString(), "THÔNG BÁO");
                return;
            }

            MessageBox.Show("XUẤT THÀNH CÔNG", "THÔNG BÁO");
        }

        private string Random_From_Now()
        {
            Random r = new Random();
            string Random_Number = r.Next(Int32.Parse(DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString()), 999999999).ToString() + " - " + r.Next(Int32.Parse(DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString()), 999999999).ToString();
            return Random_Number;
        }

        private string Process_Name_File(string Name_File)
        {
            string List_Char_Not_Allow = @"\/:*?<>|" + Convert.ToChar(34);
            for (int i = 0; i < List_Char_Not_Allow.Length; i++)
            {
                Name_File = Name_File.Replace(List_Char_Not_Allow[i].ToString().Trim(), "");
            }
            return Name_File;
        }

        private string Get_File_Ext_From_String(string File_Name)
        {
            string File_Ext = "";
            int Last_Index_Dot = File_Name.LastIndexOf(".");
            if (Last_Index_Dot > -1) { File_Ext = File_Name.Substring(Last_Index_Dot); }
            return File_Ext;
        }

        private void lst_sheet_from_excel_file_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.A && e.Control)
            {
                for (int i = 0; i < lst_sheet_from_excel_file.Items.Count; i++)
                {
                    lst_sheet_from_excel_file.SetSelected(i,true);
                }
            }
        }
    }
}
