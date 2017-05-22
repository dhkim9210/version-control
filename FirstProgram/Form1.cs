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
