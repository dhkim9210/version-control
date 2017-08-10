using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace FirstProgram
{
    public partial class Form1 : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        int num = 0;
        public static string PathA = null;
        public Form1(string a)
        {
            InitializeComponent();
            PathA = a;
        }
        public Form1()
        {
            InitializeComponent();
            
        }
        private void btnLoad_Click(object sender, EventArgs e)
        {
            //수정된 부분 : 필요없는 Sheet 제거
            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook book =
            xlApp.Workbooks.Open(PathA);
            xlApp.DisplayAlerts = false;
            Excel.Worksheet worksheet = null;

            // sheet 제거 부분

            worksheet = (Excel.Worksheet)book.Worksheets[1];
            worksheet.Delete();
            worksheet = (Excel.Worksheet)book.Worksheets[2];
            worksheet.Delete();
            worksheet = (Excel.Worksheet)book.Worksheets[2];
            worksheet.Delete();
            worksheet = (Excel.Worksheet)book.Worksheets[1];
            worksheet.Delete();
            // 가구 시작
            for (int i = 0; i < 2; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[2];
                worksheet.Delete();     // 거터 제거
            }
            // 가구, 경사로,계단 출력
            for (int i = 0; i < 11; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[4];
                worksheet.Delete();
            }

            //가구, 경사로, 계단, 구조기둥, 구조 기초 출력
            for (int i = 0; i < 2; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[7];
                worksheet.Delete();             //구조 보 시스템 태그 제거
            }
            for (int i = 0; i < 2; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[8];
                worksheet.Delete();             //그리드
            }
            //난간까지 출력
            for (int i = 0; i < 5; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[10];
                worksheet.Delete();             //그리드 헤드 제거
            }
            for (int i = 0; i < 7; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[12];
                worksheet.Delete();             // 문 출력
            }

            worksheet = (Excel.Worksheet)book.Worksheets[13];
            worksheet.Delete();             // 바닥 출력

            for (int i = 0; i < 2; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[16];
                worksheet.Delete();             // 배선,벽 출력
            }
            for (int i = 0; i < 4; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[17];
                worksheet.Delete();             // 
            }
            for (int i = 0; i < 4; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[18];
                worksheet.Delete();             // 
            }
            for (int i = 0; i < 3; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[19];
                worksheet.Delete();             // 
            }
            for (int i = 0; i < 10; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[20];
                worksheet.Delete();             // 
            }
            for (int i = 0; i < 14; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[21];
                worksheet.Delete();             // 
            }
            worksheet = (Excel.Worksheet)book.Worksheets[22];
            worksheet.Delete();             // 
            for (int i = 0; i < 3; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[24];
                worksheet.Delete();             // 
            }
            for (int i = 0; i < 8; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[25];
                worksheet.Delete();             // 
            }
            for (int i = 0; i < 2; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[26];
                worksheet.Delete();             // 
            }
            for (int i = 0; i < 2; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[27];
                worksheet.Delete();             // 
            }
            for (int i = 0; i < 2; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[32];
                worksheet.Delete();              
            }

            worksheet = (Excel.Worksheet)book.Worksheets[33];
            worksheet.Delete();              
            for (int i = 0; i < 6; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[34];
                worksheet.Delete();             
            }
            for (int i = 0; i < 2; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[35];
                worksheet.Delete();             
            }

            for (int i = 0; i < 3; i++)
            {
                worksheet = (Excel.Worksheet)book.Worksheets[38];
                worksheet.Delete();             
            }




            xlApp.DisplayAlerts = true;
            book.Save();
            book.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("FILE Save Okay : Sheet remove");

            string filePath = PathA;
            string fileExtension = Path.GetExtension(filePath);
            string header = rbHeaderYes.Checked ? "Yes" : "No";
            string connectionString = string.Empty;
            string sheetName = string.Empty;
            string a = "$";
            string comboboxitem;
            
            switch(fileExtension)
            {
                case ".xls":   
                    connectionString = string.Format(Excel03ConString, filePath,header);
                    break;
                case ".xlsx": 
                    connectionString = string.Format(Excel07ConString, filePath,header);
                    break;
            }

            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[num]["TABLE_NAME"].ToString();
                    for (int i = 0; i < dtExcelSchema.Rows.Count; i++)
                    {
                        comboboxitem = (dtExcelSchema.Rows[i]["TABLE_NAME"].ToString()).Substring(0, dtExcelSchema.Rows[i]["TABLE_NAME"].ToString().IndexOf(a));
                        if (comboboxitem.Substring(0, 1) == "'")
                        {
                            comboboxitem = (comboboxitem.Substring(1));
                        }
                        if (i == 0)
                        {
                            comboBox1.Text = comboboxitem;
                        }
                        comboBox1.Items.Add(comboboxitem);
                    }
                    con.Close();
                }
            }
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();
                        dataGridView1.AllowUserToAddRows = false; 
                        dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        dataGridView1.AutoResizeColumns();
                        dataGridView1.DataSource = dt;
                    }
                }
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            num = comboBox1.SelectedIndex;
            dataGridView1.Columns.Clear();
            comboBox1.Text = comboBox1.SelectedItem.ToString();
            string filePath = PathA;
            string fileExtension = Path.GetExtension(filePath);
            string header = rbHeaderYes.Checked ? "Yes" : "No";
            string connectionString = string.Empty;
            string sheetName = string.Empty;
            switch (fileExtension)
            {
                case ".xls":
                    connectionString = string.Format(Excel03ConString, filePath, header);
                    break;
                case ".xlsx":
                    connectionString = string.Format(Excel07ConString, filePath, header);
                    break;
            }
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[num]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();
                        dataGridView1.DataSource = dt;
                    }
                }
            }

        }
    }
}
