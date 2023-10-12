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
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportDSKH
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void DataTableSetUp()
        {
            DataTable table = new DataTable();
        }
        private void ImportDuLieuExcel(string path)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];
                //DataTable dt = new DataTable();
                //for(int i = excelWorksheet.Dimension.Start.Column; i <= excelWorksheet.Dimension.End.Column - 1; i++)
                //{
                //dt.Columns.Add("TT");
                //dt.Columns.Add("Mã KH");
                //dt.Columns.Add("Tên KH");
                //dt.Columns.Add("Ngày sinh");
                //dt.Columns.Add("SDT");
                //dt.Columns.Add("Email");
                //dt.Columns.Add("Địa chỉ");
                    //dt.Columns.Add(excelWorksheet.Cells[1,i].Value.ToString());
                //}
                for(int i2 = excelWorksheet.Dimension.Start.Row + 1; i2 <=  excelWorksheet.Dimension.End.Row; i2++) 
                {
                    // List<string> listRows = new List<string>();
                    //listRows.Add("");
                    int index = dataGridView1.Rows.Add();
                    for (int j = excelWorksheet.Dimension.Start.Column; j < excelWorksheet.Dimension.End.Column; j++)
                    {
                        //listRows.Add(excelWorksheet.Cells[i2,j].Value.ToString());
                        dataGridView1.Rows[index].Cells[j].Value = excelWorksheet.Cells[i2, j].Value.ToString();
                    }
                    //dt.Rows.Add(listRows.ToArray());
                }
                //dataGridView1.DataSource = dt;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = " Import Excel ";
            openFileDialog.Filter = "Excel (*.xlsx) | *.xlsx";
            if(openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ImportDuLieuExcel(openFileDialog.FileName);
                textBox1.Text = openFileDialog.FileName;
            }
        }
        private int KTKhachHang()
        {
            for(int i = 1; i <= dataGridView1.Rows.Count; i++)
            {
                if(dataGridView1.Rows[i].Cells[1].Value.ToString() == textBox2.Text)
                {
                    return i;
                }
            }
            return -1;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            int Check = KTKhachHang();
            if(Check != -1) 
            {
                dataGridView1.Rows[Check].Cells[4].Value = textBox3.Text;
            }
        }
    }
}
