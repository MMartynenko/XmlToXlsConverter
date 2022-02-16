﻿using System;
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
using System.Xml;

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
                    XmlDocument dt = CreateDataTableFromXml(txtXmlFilePath.Text);
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
                    XmlDocument dt = CreateDataTableFromXml(txtXmlFilePath.Text);
                    ExportDataTableToExcel(dt, XlFile);

                    MessageBox.Show("Conversion completed");
                }
            }
            else
            {
                MessageBox.Show("Please fill required fields");
            }
        }

        public XmlDocument CreateDataTableFromXml(string XmlFile)
        {          
            
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(XmlFile);
                return doc;                

            }
            catch (Exception ex)
            {

            }
            return null;
        }

        private void ExportDataTableToExcel(XmlDocument doc, string Xlfile)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook book = excel.Application.Workbooks.Add(Type.Missing);
            excel.Visible = false;
            excel.DisplayAlerts = false;

            XmlNodeList worksheets = doc.GetElementsByTagName("Worksheet");
            XmlNodeList styles = doc.GetElementsByTagName("Styles");

            FillStyles(book, styles);
            FillContent(book, worksheets);

            book.SaveAs(Xlfile);
            book.Close(true);
            excel.Quit();

            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(excel);
        }

        private void FillStyles(Workbook book, XmlNodeList styles)
        {
            if (styles == null || styles.Count < 1) return;

            Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet;
            List<XmlNode> stylesList = new List<XmlNode>();
            XmlNodeList children = styles[0].ChildNodes;
            foreach (XmlNode child in children)
            {
                if (child.Name == "Style") stylesList.Add(child);
            }

            foreach (XmlNode style in stylesList)
            {
                string id = style.Attributes["ss:ID"].Value;
                Style newStyle = book.Styles.Add(id);

                XmlNodeList parts = style.ChildNodes;
                foreach (XmlNode part in parts)
                {                    
                    switch (part.Name)
                    {
                        case "Alignment":
                            if (part.Attributes["ss:Horizontal"] != null) {
                                XlHAlign alignH;
                                if (Enum.TryParse("xlHAlign" + part.Attributes["ss:Horizontal"].Value, out alignH)) newStyle.HorizontalAlignment = alignH;                            
                            }
                            if (part.Attributes["ss:Vertical"] != null)
                            {
                                XlVAlign alignV;
                                if (Enum.TryParse("xlVAlign" + part.Attributes["ss:Vertical"].Value, out alignV)) newStyle.VerticalAlignment = alignV;
                            }
                            if (part.Attributes["ss:Indent"] != null) newStyle.IndentLevel = Int32.Parse(part.Attributes["ss:Indent"].Value);
                            if (part.Attributes["ss:ShrinkToFit"] != null) newStyle.ShrinkToFit = part.Attributes["ss:ShrinkToFit"].Value == "1";
                            if (part.Attributes["ss:WrapText"] != null) newStyle.WrapText = part.Attributes["ss:WrapText"].Value == "1";
                            break;                        
                        case "Font":
                            if (part.Attributes["ss:Bold"] != null) newStyle.Font.Bold = part.Attributes["ss:Bold"].Value == "1";
                            if (part.Attributes["ss:Color"] != null) newStyle.Font.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(part.Attributes["ss:Color"].Value));
                            if (part.Attributes["ss:FontName"] != null) newStyle.Font.Name = part.Attributes["ss:FontName"].Value;
                            if (part.Attributes["ss:Italic"] != null) newStyle.Font.Italic = part.Attributes["ss:Italic"].Value == "1";
                            if (part.Attributes["ss:Outline"] != null) newStyle.Font.OutlineFont = part.Attributes["ss:Outline"].Value == "1";
                            if (part.Attributes["ss:Shadow"] != null) newStyle.Font.Shadow = part.Attributes["ss:Shadow"].Value == "1";
                            if (part.Attributes["ss:Size"] != null) newStyle.Font.Size = Double.Parse(part.Attributes["ss:Size"].Value);
                            if (part.Attributes["ss:StrikeThrough"] != null) newStyle.Font.Strikethrough = part.Attributes["ss:StrikeThrough"].Value == "1";
                            if (part.Attributes["ss:Underline"] != null)
                            {
                                XlUnderlineStyle underline;
                                if (Enum.TryParse("xlUnderlineStyle" + part.Attributes["ss:Underline"].Value, out underline)) newStyle.Font.Underline = underline;
                            }
                            break;
                        case "Interior":
                            if (part.Attributes["ss:Color"] != null) newStyle.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(part.Attributes["ss:Color"].Value));
                            if (part.Attributes["ss:Pattern"] != null)
                            {
                                XlPattern pattern;
                                if (Enum.TryParse("xlPattern" + part.Attributes["ss:Pattern"].Value, out pattern)) newStyle.Interior.Pattern = pattern;
                            }
                            if (part.Attributes["ss:PatternColor"] != null) newStyle.Interior.PatternColor = ColorTranslator.ToOle(ColorTranslator.FromHtml(part.Attributes["ss:PatternColor"].Value));
                            break;
                        case "NumberFormat":
                            if (part.Attributes["ss:Format"] != null) newStyle.NumberFormat = part.Attributes["ss:Format"].Value;
                            break;
                        case "Protection":
                            if (part.Attributes["ss:Protected"] != null) newStyle.Locked = part.Attributes["ss:Protected"].Value == "1";
                            if (part.Attributes["ss:HideFormula"] != null) newStyle.FormulaHidden = part.Attributes["ss:HideFormula"].Value == "1";
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private void FillContent(Workbook book, XmlNodeList worksheets)
        {
            for (int i = 0; i < worksheets.Count; i++)
            {
                Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet;
                excelWorkSheet.Name = worksheets[i].Attributes["ss:Name"].Value;

                List<XmlNode> rows = new List<XmlNode>();
                List<XmlNode> columns = new List<XmlNode>();
                XmlNodeList children = worksheets[i].FirstChild.ChildNodes;
                foreach (XmlNode child in children)
                {
                    if (child.Name == "Column") columns.Add(child);
                    if (child.Name == "Row") rows.Add(child);
                }

                int columnIndex = 1;
                foreach (XmlNode column in columns)
                {
                    int span = 1;
                    if (column.Attributes["ss:Span"] != null) span += Int32.Parse(column.Attributes["ss:Span"].Value);
                    if (column.Attributes["ss:Index"] != null) columnIndex = Int32.Parse(column.Attributes["ss:Index"].Value);
                    for (int s = 0; s < span; s++)
                    {                        
                        if (column.Attributes["ss:Hidden"] != null) excelWorkSheet.Columns[columnIndex].Hidden = column.Attributes["ss:Hidden"].Value == "1";
                        if (column.Attributes["ss:Width"] != null) excelWorkSheet.Columns[columnIndex].ColumnWidth = Double.Parse(column.Attributes["ss:Width"].Value) / 5.7d;
                        if (column.Attributes["ss:StyleID"] != null) excelWorkSheet.Columns[columnIndex].Style = book.Styles[column.Attributes["ss:StyleID"].Value];
                            columnIndex++;
                    }
                    
                }

                progressBar1.Maximum = rows.Count;
                int rowIndex = 1;
                for (int j = 0; j < rows.Count; j++)
                {
                    XmlNode row = rows[j];                    
                    int span = 1;

                    if (row.Attributes["ss:Span"] != null) span += Int32.Parse(row.Attributes["ss:Span"].Value);
                    if (row.Attributes["ss:Index"] != null) rowIndex = Int32.Parse(row.Attributes["ss:Index"].Value);

                    for (int s = 0; s < span; s++)
                    {
                        if (row.Attributes["ss:Hidden"] != null) excelWorkSheet.Rows[rowIndex].Hidden = row.Attributes["ss:Hidden"].Value == "1";
                        if (row.Attributes["ss:Height"] != null) excelWorkSheet.Rows[rowIndex].RowHeight = Double.Parse(row.Attributes["ss:Height"].Value);
                        if (row.Attributes["ss:StyleID"] != null) excelWorkSheet.Rows[rowIndex].Style = book.Styles[row.Attributes["ss:StyleID"].Value];

                        int column = 1;
                        foreach (XmlNode c in row.ChildNodes)
                        {
                            if (c.Name == "Cell")
                            {
                                if (c.Attributes["ss:Index"] != null) column = Int32.Parse(c.Attributes["ss:Index"].Value);
                                if (c.Attributes["ss:StyleID"] != null) excelWorkSheet.Cells[rowIndex, column].Style = book.Styles[c.Attributes["ss:StyleID"].Value];

                                string innerText = "";
                                XmlNode data = c.FirstChild;
                                if (data.Name == "Data")
                                {
                                    if (data.Attributes["ss:Type"] != null && data.Attributes["x:Ticked"] != null && data.Attributes["ss:Type"].Value == "String" && data.Attributes["x:Ticked"].Value == "1") innerText = "'";
                                }
                                innerText += c.InnerText;
                                
                                excelWorkSheet.Cells[rowIndex, column] = innerText;
                                column++;
                                if (progressBar1.Value < progressBar1.Maximum)
                                {
                                    progressBar1.Value++;
                                    int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
                                    progressBar1.CreateGraphics().DrawString(percent.ToString() + "%", new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            }
                        }
                        rowIndex++;
                    }
                }

                if (i < worksheets.Count - 1)
                {
                    book.Worksheets.Add();
                }
            }
        }
        
    }
}
