using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace XmlToXlsConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult drResult = OFD.ShowDialog();
            if (drResult == System.Windows.Forms.DialogResult.OK)
                txtXmlFilePath.Text = OFD.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            if (chkCustomName.Checked && txtCustomFileName.Text != "" && txtXmlFilePath.Text != "") // using Custome Xml File Name  
            {
                if (File.Exists(txtXmlFilePath.Text))
                {
                    string CustXmlFilePath = Path.Combine(new FileInfo(txtXmlFilePath.Text).DirectoryName, txtCustomFileName.Text); // Ceating Path for Xml Files  
                    System.Data.DataTable dt = CreateDataTableFromXml(txtXmlFilePath.Text);
                    ExportDataTableToExcel(dt, CustXmlFilePath);

                    MessageBox.Show("Conversion completed");
                }

            }
            else if (!chkCustomName.Checked || txtXmlFilePath.Text != "") // Using Default Xml File Name  
            {
                if (File.Exists(txtXmlFilePath.Text))
                {
                    FileInfo fi = new FileInfo(txtXmlFilePath.Text);
                    string XlFile = fi.DirectoryName + "\\" + fi.Name.Replace(fi.Extension, ".xlsx");
                    System.Data.DataTable dt = CreateDataTableFromXml(txtXmlFilePath.Text);
                    ExportDataTableToExcel(dt, XlFile);

                    MessageBox.Show("Conversion completed");
                }
            }
            else
            {
                MessageBox.Show("Please fill required fields");
            }
        }

        public System.Data.DataTable CreateDataTableFromXml(string XmlFile)
        {

            System.Data.DataTable Dt = new System.Data.DataTable();
            try
            {
                DataSet ds = new DataSet();
                ds.ReadXml(XmlFile);
                Dt.Load(ds.CreateDataReader());

            }
            catch (Exception ex)
            {

            }
            return Dt;
        }

        private void ExportDataTableToExcel(System.Data.DataTable table, string Xlfile)
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook book = excel.Application.Workbooks.Add(Type.Missing);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet;
            excelWorkSheet.Name = table.TableName;

            progressBar1.Maximum = table.Columns.Count;
            for (int i = 1; i < table.Columns.Count + 1; i++) // Creating Header Column In Excel  
            {
                excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                if (progressBar1.Value < progressBar1.Maximum)
                {
                    progressBar1.Value++;
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString(percent.ToString() + "%", new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
                    System.Windows.Forms.Application.DoEvents();
                }
            }


            progressBar1.Maximum = table.Rows.Count;
            for (int j = 0; j < table.Rows.Count; j++) // Exporting Rows in Excel  
            {
                for (int k = 0; k < table.Columns.Count; k++)
                {
                    excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                }

                if (progressBar1.Value < progressBar1.Maximum)
                {
                    progressBar1.Value++;
                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                    progressBar1.CreateGraphics().DrawString(percent.ToString() + "%", new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
                    System.Windows.Forms.Application.DoEvents();
                }
            }


            book.SaveAs(Xlfile);
            book.Close(true);
            excel.Quit();

            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(excel);

        }
    }
}
