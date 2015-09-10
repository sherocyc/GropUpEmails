﻿using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using GropUpEmails.Properties;
using Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;

namespace GropUpEmails
{
    public partial class GroupUpEmails : Form
    {
        private const string STYLE_TITLE =
            "HEIGHT: 23.25pt; " +
            "WIDTH: 92pt; " +
            "BORDER-TOP: windowtext 0.5pt solid;" +
            "BORDER-RIGHT: windowtext 0.5pt solid; " +
            "BORDER-BOTTOM: windowtext 0.5pt solid; " +
            "BORDER-LEFT: windowtext 0.5pt solid; " +
            "BACKGROUND-COLOR: yellow";
        public GroupUpEmails()
        {
            InitializeComponent();
            btnReciever.Click += (sender, e) =>{
                GetRecieverFilePath();
                GetRecievers(txtRecieverFile.Text);
            };
            btnData.Click += (sender, e) => {
                GetDataFilePath();
                GenerateDataPreview(txtDataFile.Text);
            };

            btnOK.Click += (sender, e) => {
                GetMailSend(recieverGridView.SelectedRows[0].Cells[3].Value.ToString());
                //GetAndWrite(txtDataFile.Text, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+"\\123.xml");
            };
            regenarateBtn.Click += (sender, e) => {
                GenerateDataPreview(txtDataFile.Text);
            };
            FormClosing += (sender, e) => {
                if (MessageBox.Show(Resources.ComfirmQuit, Resources.Quit, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    e.Cancel = true;
                    WindowState = FormWindowState.Minimized;
                    ShowInTaskbar = false;
                    Application.Exit();
                }
                else {
                    e.Cancel = true;
                }
            };
            
        }
        private void GetRecievers(string xlsFilePath) {
            try
            {
                progressBar.Value = 0;
                string strConn = $"Provider=Microsoft.Ace.OleDb.12.0;Data Source={xlsFilePath};Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";
                string strComm = "SELECT 教师姓名,教师证件号,邮箱号 FROM [Sheet1$]";
                OleDbConnection myConn = new OleDbConnection(strConn);
                OleDbDataAdapter myAdp = new OleDbDataAdapter(strComm, strConn);
                DataSet ds = new DataSet();
                myAdp.Fill(ds);
                myConn.Close();
                progressBar.Value = 50;

                DataGridViewCheckBoxColumn Column1 = new DataGridViewCheckBoxColumn
                {
                    HeaderText = "选择",
                    Name = "选择",
                    ReadOnly = false
                };
                recieverGridView.Columns.Clear();
                recieverGridView.Columns.Add(Column1);
                recieverGridView.DataSource = ds.Tables[0];
                recieverGridView.CurrentCell = recieverGridView.Rows[0].Cells[0];
                recieverGridView.Columns[0].Width = 30;
                foreach (DataGridViewRow row in recieverGridView.Rows)
                {
                    row.Cells[0].Value = true;
                }
                progressBar.Value = 100;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }
        private void GetRecieverFilePath() {
            OpenFileDialog openFileDialog = new OpenFileDialog {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Filter = Resources.ExcelFilter, 
                FilterIndex = 1,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK) return;
            try {
                Stream myStream = openFileDialog.OpenFile();
                using (myStream) {
                }
                txtRecieverFile.Text = openFileDialog.FileName;
            }
            catch (Exception e) {
                MessageBox.Show(e.Message);
            }
        }
        private void GetDataFilePath() {
            OpenFileDialog openFileDialog = new OpenFileDialog {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Filter = Resources.ExcelFilter,
                FilterIndex = 1,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;
            try {
                Stream myStream = openFileDialog.OpenFile();
                using (myStream) {
                }
                txtDataFile.Text = openFileDialog.FileName;
            }
            catch (Exception e) {
                MessageBox.Show(e.Message);
            }
        }
        private void GenerateDataPreview(string xlsFilePath)
        {
            contentEditor.Text = "";
            progressBar.Value = 0;
            try
            {
                progressBar.Value = 1;
                string strConn = $"Provider=Microsoft.Ace.OleDb.12.0;Data Source={xlsFilePath};Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";
                string strComm = $"SELECT 教师姓名,教师证件号,年级,科目,实际单价,课时,学生姓名,班主任,教学点 FROM [Sheet1$] WHERE 教师证件号 = \"{recieverGridView.SelectedRows[0].Cells[2].Value}\" ";
                OleDbConnection myConn = new OleDbConnection(strConn);
                OleDbDataAdapter myAdp = new OleDbDataAdapter(strComm, strConn);
                DataSet ds = new DataSet();
                myAdp.Fill(ds);
                myConn.Close();

                txtTitle.Text = Convert.ToString(ds.Tables[0].Rows[0][0]);
                contentEditor.Text = GenerateContent(ds.Tables[0]).OuterXml;

                progressBar.Value = 100;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void SendTo(SmtpClient client, string recieverEmail, string title, string content) {
            try
            {
                MailMessage message = new MailMessage(txtSender.Text, recieverEmail) {
                    Subject = title,
                    Body = content,
                    BodyEncoding = Encoding.UTF8,
                    IsBodyHtml = true
                };
                client.Send(message);
                MessageBox.Show("发送邮件成功！");
            }
            catch (Exception e) {
                MessageBox.Show(e.Message);
            }
        }
        protected void GetMailSend(string reciever)
        {
            try
            {
                txtSender.Text = txtSender.Text.TrimEnd("@qq.com".ToCharArray()) + @"@qq.com";
                SmtpClient client = new SmtpClient {
                    Host = "smtp.qq.com",
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(txtSender.Text, txtPwd.Text),
                    DeliveryMethod = SmtpDeliveryMethod.Network
                };
                SendTo(client,reciever,txtTitle.Text,contentEditor.Text);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        private XmlAttribute createAttribute(XmlDocument doc, string name, string value)
        {
            XmlAttribute attr = doc.CreateAttribute(name);
            attr.Value = value;
            return attr;
        }
        private XmlDocument GenerateContent(DataTable data) {
            try
            {
                XmlDocument xml = new XmlDocument();
                XmlNode div = xml.CreateElement("div");
                XmlNode table = xml.CreateElement("table");
                table.Attributes.Append(createAttribute(xml,"style", "WIDTH: 919pt; BORDER-COLLAPSE: collapse"));
                table.Attributes.Append(createAttribute(xml, "cellspacing", "0"));
                table.Attributes.Append(createAttribute(xml, "cellpadding", "0"));
                table.Attributes.Append(createAttribute(xml, "width", "1225"));
                table.Attributes.Append(createAttribute(xml, "border", "0"));
                XmlNode colgroup = xml.CreateElement("colgroup");
                foreach (var column in data.Columns)
                {
                    XmlNode col = xml.CreateElement("col");
                    col.Attributes.Append(createAttribute(xml, "style", "WIDTH: 92pt; mso-width-source: userset; mso-width-alt: 4380"));
                    col.Attributes.Append(createAttribute(xml, "width", "123"));
                    colgroup.AppendChild(col);
                }
                XmlNode tbody = xml.CreateElement("tbody");
                
                {
                    XmlNode tr = xml.CreateElement("tr");
                    tr.Attributes.Append(createAttribute(xml, "style", "HEIGHT: 23.25pt; mso-height-source: userset"));
                    tr.Attributes.Append(createAttribute(xml, "height", "30"));
                    tbody.AppendChild(tr);
                    foreach (var column in data.Columns) {
                        XmlNode td = xml.CreateElement("td");
                        td.Attributes.Append(createAttribute(xml, "class", "xl23212"));
                        td.Attributes.Append(createAttribute(xml, "style", STYLE_TITLE));
                        td.Attributes.Append(createAttribute(xml, "height", "30"));
                        td.Attributes.Append(createAttribute(xml, "width", "123"));
                        XmlNode strong = xml.CreateElement("strong");
                        XmlNode font = xml.CreateElement("font");
                        font.Attributes.Append(createAttribute(xml, "size", "2"));
                        font.Attributes.Append(createAttribute(xml, "face", "宋体"));
                        font.InnerText = column.ToString();
                        strong.AppendChild(font);
                        td.AppendChild(strong);
                        tbody.AppendChild(td);
                    }
                }

                foreach (DataRow row in data.Rows)
                {
                    XmlNode tr = xml.CreateElement("tr");
                    tr.Attributes.Append(createAttribute(xml, "style", "HEIGHT: 12pt"));
                    tr.Attributes.Append(createAttribute(xml, "height", "16"));
                    tbody.AppendChild(tr);
                    for(int i=0;i<data.Columns.Count;i++)
                    {
                        XmlNode td = xml.CreateElement("td");
                        td.Attributes.Append(createAttribute(xml, "class", "xl23212"));
                        td.Attributes.Append(createAttribute(xml, "style", "BORDER-TOP: windowtext; HEIGHT: 12pt; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"));
                        td.Attributes.Append(createAttribute(xml, "height", "16"));
                        td.Attributes.Append(createAttribute(xml, "width", "123"));
                        XmlNode font = xml.CreateElement("font");
                        font.Attributes.Append(createAttribute(xml, "size", "2"));
                        font.Attributes.Append(createAttribute(xml, "face", "宋体"));
                        font.InnerText = row[i].ToString();
                        td.AppendChild(font);
                        tbody.AppendChild(td);
                    }
                }
                table.AppendChild(tbody);
                table.AppendChild(colgroup);
                div.AppendChild(table);
                xml.AppendChild(div);
                return xml;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }


            /*
            content = table.Columns.Cast<DataColumn>().Aggregate("", (current, dc) => current + dc.ColumnName.PadRight(10));
            txtContent.Text += Resources.Endline;
            foreach (DataRow row in table.Rows) {
                for (int i = 0; i < table.Columns.Count; i++)
                    content += row[i].ToString().PadRight(10);
                content += Resources.Endline;
            }*/
        }

        public void GetAndWrite(string strFileName, string strDesPath) {
            Microsoft.Office.Interop.Excel.Application excelRead = new ApplicationClass();
            Workbooks workbooks = excelRead.Workbooks;
            object mo = Missing.Value;
            Workbook workbook = workbooks.Open(strFileName, mo, mo, mo, mo, mo, mo, mo, mo, mo, mo, mo, mo, mo, mo);
            string[] strSheetNames = {"Sheet1$"};

            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];
            int bgnRow = (worksheet.UsedRange.Cells.Row > 1) ? worksheet.UsedRange.Cells.Row - 1 : worksheet.UsedRange.Cells.Row;
            int bgnColumn = (worksheet.UsedRange.Cells.Column > 1) ? worksheet.UsedRange.Cells.Column - 1 : worksheet.UsedRange.Cells.Column;
            //设置保存xml文件
            string strFile = strDesPath.Substring(strDesPath.LastIndexOf("\\", StringComparison.Ordinal) + 1);
            strFile = strFile.Remove(strFile.IndexOf(".", StringComparison.Ordinal), 4);

            string strPath = strDesPath.Substring(0, strDesPath.LastIndexOf("\\", StringComparison.Ordinal));
            string fileName = strPath + "\\" + strFile + ".xml";
            
            XmlTextWriter writexml = new XmlTextWriter(fileName, Encoding.Default);
            writexml.Formatting = Formatting.Indented;
            writexml.WriteStartDocument();
            writexml.WriteStartElement("workbook");
            writexml.WriteStartElement("worksheet");
            writexml.WriteAttributeString("name", strSheetNames[0]);
            writexml.WriteStartElement("table");
            writexml.WriteAttributeString("sumRow", worksheet.UsedRange.Cells.Rows.Count.ToString());
            writexml.WriteAttributeString("sumCol", worksheet.UsedRange.Cells.Columns.Count.ToString());

            for (int i = bgnRow; i < worksheet.UsedRange.Cells.Rows.Count + bgnRow; ++i) {
                writexml.WriteStartElement("row");
                for (int j = bgnColumn; j < worksheet.UsedRange.Cells.Columns.Count + bgnColumn; ++j) {
                    string strData = ((Range)worksheet.UsedRange.Cells[i, j]).Text.ToString();
                    int mergeAcross = ((Range)worksheet.Cells[i, j]).MergeArea.Count;
                    mergeAcross -= 1;
                    string backgroundColor = ((Range)worksheet.Cells[i, j]).Interior.Color.ToString();
                    Int32 bgcolor = Int32.Parse(backgroundColor);
                    Color backClor = Color.FromArgb(bgcolor);
                    string Bgcolor = ColorTranslator.ToHtml(backClor);

                    string fontFamily = ((Range)worksheet.Cells[i, j]).Font.Name.ToString();

                    string fontColor = ((Range)worksheet.Cells[i, j]).Font.Color.ToString();
                    int fcolor = int.Parse(fontColor);
                    Color FontColor = Color.FromArgb(fcolor);
                    string font_color = ColorTranslator.ToHtml(FontColor);

                    string fontSize = ((Range)worksheet.Cells[i, j]).Font.Size.ToString();
                    if (mergeAcross > 1) {
                        j += mergeAcross;
                    }
                    writexml.WriteStartElement("cell");
                    writexml.WriteAttributeString("isMerge", mergeAcross.ToString());
                    if ((backgroundColor == "16777215") && (Equals(fontFamily, "宋体")) && (fontColor == "0") && (fontSize == "12")) {
                        writexml.WriteAttributeString("style", "");
                    }
                    else {
                        writexml.WriteAttributeString("style", " background-color:" + Bgcolor + ";" + "font-family:" + fontFamily + ";" + "font-color:" + font_color + ";" + "font-size:" + fontSize + ";");
                    }
                    writexml.WriteString(strData);
                    writexml.WriteEndElement();
                }
                writexml.WriteEndElement();
            }
            writexml.WriteEndElement();
            writexml.WriteEndElement();
            writexml.WriteEndElement();
            writexml.WriteEndDocument();
            writexml.Close();
            excelRead.Quit();
        }
    }
}

