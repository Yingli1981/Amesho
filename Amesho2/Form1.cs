using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Amesho2
{
    public partial class FormMain : Form
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public FormMain()
        {
            InitializeComponent();

            // DataGridView1にユーザーが新しい行を追加できないようにする
            dataGridView1.AllowUserToAddRows = false;
        }

        /// <summary>
        /// [参照]ボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "Excelファイル|*.xls;*.xlsx;*.xlsm";
            ofd.Title = "編集対象のExcelファイルを選択してください。";
            ofd.RestoreDirectory = true;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtExcelFile.Text = ofd.FileName;
            }

            // Excelファイル展開
            ShowHeaderFooter(txtExcelFile.Text);
        }

        /// <summary>
        /// Excelファイル展開し、グリッドに表示
        /// </summary>
        /// <param name="strPath">Excelファイル名(フルパス)</param>
        private void ShowHeaderFooter(string strPath)
        {
            if (strPath == "") { return; }

            // 待機状態
            Cursor.Current = Cursors.WaitCursor;

            dataGridView1.Rows.Clear();
            
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            // エクセルを非表示
            ExcelApp.Visible = false;

            // エクセルファイルのオープン
            Microsoft.Office.Interop.Excel.Workbook WorkBook = ExcelApp.Workbooks.Open(strPath);

            int iRowCnt = 0;
            for (int i = 0; i < WorkBook.Sheets.Count; i++)
            {
                // シートの選択
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)WorkBook.Sheets[(i + 1)];

                if ((sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden)
                 || (sheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVeryHidden))
                {
                    continue;
                }

                sheet.Select();

                dataGridView1.Rows.Add();

                dataGridView1.Rows[iRowCnt].Cells["colSheetName"].Value = sheet.Name;
                dataGridView1.Rows[iRowCnt].Cells["colLHeader"].Value = sheet.PageSetup.LeftHeader;
                dataGridView1.Rows[iRowCnt].Cells["colCHeader"].Value = sheet.PageSetup.CenterHeader;
                dataGridView1.Rows[iRowCnt].Cells["colRHeader"].Value = sheet.PageSetup.RightHeader;
                dataGridView1.Rows[iRowCnt].Cells["colLFooter"].Value = sheet.PageSetup.LeftFooter;
                dataGridView1.Rows[iRowCnt].Cells["colCFooter"].Value = sheet.PageSetup.CenterFooter;
                dataGridView1.Rows[iRowCnt].Cells["colRFooter"].Value = sheet.PageSetup.RightFooter;
                iRowCnt++;
            }

            // workbookを閉じる
            try 
            {
                WorkBook.Close(SaveChanges: false);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(WorkBook);
            }

            // エクセルを閉じる
            try 
            {
                ExcelApp.Quit();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
            }
            // 元に戻す
            Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// [設定変更開始]ボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            // 待機状態
            Cursor.Current = Cursors.WaitCursor;

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            // エクセルを非表示
            ExcelApp.Visible = false;

            // エクセルファイルのオープン
            Microsoft.Office.Interop.Excel.Workbook WorkBook = ExcelApp.Workbooks.Open(txtExcelFile.Text);

            int i = 0;

            label4.Text = i.ToString() + " / " + dataGridView1.Rows.Count.ToString();

            for ( ; i < dataGridView1.Rows.Count; i++)
            {
                string strSheetName = dataGridView1.Rows[i].Cells["colSheetName"].Value.ToString();

                // シートの選択
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)WorkBook.Worksheets[strSheetName];
                sheet.Select();

                sheet.PageSetup.LeftHeader = Convert.ToString(dataGridView1.Rows[i].Cells["colLHeader"].Value);
                sheet.PageSetup.CenterHeader = Convert.ToString(dataGridView1.Rows[i].Cells["colCHeader"].Value);
                sheet.PageSetup.RightHeader = Convert.ToString(dataGridView1.Rows[i].Cells["colRHeader"].Value);
                sheet.PageSetup.LeftFooter = Convert.ToString(dataGridView1.Rows[i].Cells["colLFooter"].Value);
                sheet.PageSetup.CenterFooter = Convert.ToString(dataGridView1.Rows[i].Cells["colCFooter"].Value);
                sheet.PageSetup.RightFooter = Convert.ToString(dataGridView1.Rows[i].Cells["colRFooter"].Value);

                label4.Text = (i+1).ToString() + " / " + dataGridView1.Rows.Count.ToString();
                label4.Update();
            }

            // 保存メッセージを出さないようにする
            ExcelApp.DisplayAlerts = false;
            
            // 保存
            WorkBook.Save();

            // workbookを閉じる
            try 
            {
                WorkBook.Close();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(WorkBook);
            }

            // エクセルを閉じる
            try 
            {
                ExcelApp.Quit();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
            }

            // 元に戻す
            Cursor.Current = Cursors.Default;


            MessageBox.Show("ヘッダー・フッターの編集が完了しました。");
        }
    }
}
