using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using GropUpEmails.Properties;
using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;

namespace GropUpEmails
{
    public partial class GroupUpEmails : Form
    {
        private GroupUpEmailManager manager;

        private const string  STYLE_HEADER =
            "HEIGHT: 23.25pt; " +
            "WIDTH: 92pt; " +
            "BORDER-TOP: windowtext 0.5pt solid;" +
            "BORDER-RIGHT: windowtext 0.5pt solid; " +
            "BORDER-BOTTOM: windowtext 0.5pt solid; " +
            "BORDER-LEFT: windowtext 0.5pt solid; " +
            "BACKGROUND-COLOR: yellow";

        private const string STYLE_CONTENT =
            "HEIGHT: 12pt; " +
            "BORDER-TOP: windowtext; " +
            "BORDER-RIGHT: windowtext 0.5pt solid; " +
            "BORDER-BOTTOM: windowtext 0.5pt solid; " +
            "BORDER-LEFT: windowtext 0.5pt solid; " +
            "BACKGROUND-COLOR: transparent";
        
        private DataTable _detailTable;
        private DataTable _calculateTable;

        public GroupUpEmails()
        {
            InitializeComponent();
            manager = new GroupUpEmailManager( this );
            txtSender.Text = UserPreference.Instance.Data.UserEmail;
            txtPwd.Text = UserPreference.Instance.Data.UserEmailPassword;
            btnData.Enabled = false;
            regenarateBtn.Enabled = false;
            btnReciever.Click += (sender, e) =>{
                Utils.OpenXlsFile(filename => {
                    txtRecieverFile.Text = filename;
                    GetRecievers( txtRecieverFile.Text );
                    btnData.Enabled = true;
                });
            };
            btnData.Click += (sender, e) => {
                Utils.OpenXlsFile( filename => {
                    txtDataFile.Text = filename;
                    GetDataTable( txtDataFile.Text , GenerateDataPreview );
                    regenarateBtn.Enabled = true;
                } );
            };

            btnOK.Click += (sender, e) => {
                GetMailSend();
            };
            regenarateBtn.Click += (sender, e) => {
                GenerateDataPreview();
            };

            FormClosing += (sender, e) => {
                if (MessageBox.Show(Resources.ComfirmQuit, Resources.Quit, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    e.Cancel = true;
                    WindowState = FormWindowState.Minimized;
                    ShowInTaskbar = false;
                    UserPreference.Instance.Data.UserEmail = txtSender.Text;
                    UserPreference.Instance.Data.UserEmailPassword = txtPwd.Text;
                    UserPreference.Instance.Data.RecieverDataPath = txtRecieverFile.Text ;
                    UserPreference.Instance.Save();
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
                string strConn = $"Provider=Microsoft.Ace.OleDb.12.0;Data Source={xlsFilePath};Extended Properties='Excel 12.0; HDR=Yes; IMEX=1;'";
                string strComm = "SELECT 教师姓名,教师证件号,邮箱号 FROM [Sheet1$]";
                OleDbConnection myConn = new OleDbConnection(strConn);
                OleDbDataAdapter myAdp = new OleDbDataAdapter(strComm, strConn);
                DataSet ds = new DataSet();
                myAdp.Fill(ds);
                myConn.Close();
                progressBar.Value = 50;

                DataGridViewCheckBoxColumn column1 = new DataGridViewCheckBoxColumn
                {
                    HeaderText = "选择",
                    Name = "选择",
                    ReadOnly = false
                };
                recieverGridView.Columns.Clear();
                recieverGridView.Columns.Add(column1);
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

        private void GetDataTable(string xlsFilePath , Action callback) {
            BackgroundWorker bgWork = new BackgroundWorker();
            bgWork.DoWork += (sender, e) => {
                try {
                    string strConn = $"Provider=Microsoft.Ace.OleDb.12.0;Data Source={xlsFilePath};Extended Properties='Excel 12.0; HDR=Yes;IMEX=1;'";
                    string strComm = "SELECT * FROM [英语$] ";
                    OleDbConnection myConn = new OleDbConnection(strConn);
                    myConn.Open();
                    OleDbDataAdapter myAdp = new OleDbDataAdapter(strComm, strConn);
                    DataSet ds = new DataSet();
                    myAdp.Fill(ds);
                    _detailTable = ds.Tables[0];

                    strComm = "SELECT * FROM [计算列$] ";
                    myAdp = new OleDbDataAdapter(strComm, strConn);
                    ds = new DataSet();
                    myAdp.Fill(ds);
                    _calculateTable = ds.Tables[0];
                    myConn.Close();
                }
                catch (Exception ex) {
                    MessageBox.Show(ex.Message);
                }
            };
            bgWork.RunWorkerCompleted += (sender, e) =>
            {
                callback?.Invoke();
            };
            bgWork.RunWorkerAsync();
        }
        private void GenerateDataPreview()
        {
            contentEditor.Text = "";
            progressBar.Value = 0;
            try
            {
                txtTitle.Text = Convert.ToString(recieverGridView.SelectedRows[0].Cells[1].Value);

                String str;
                DataRow[] results = _detailTable.Select($"教师姓名 = '{recieverGridView.SelectedRows[0].Cells[1].Value}' ");
                DataTable t = _detailTable.Clone();
                foreach (DataRow row in results)
                {
                    t.ImportRow(row);
                }
                str = GenerateContent( t ).OuterXml;

                results = _calculateTable.Select( $"教师姓名 = '{recieverGridView.SelectedRows[0].Cells[1].Value}' " );
                t = _calculateTable.Clone();
                foreach ( DataRow row in results ) {
                    t.ImportRow( row );
                }
                str += GenerateContent( t ).OuterXml;

                contentEditor.Text = str;
                progressBar.Value = 100;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void SendTo(SmtpClient client, SendEmailModel model) {
            MailMessage message = new MailMessage(txtSender.Text, model.recieverEmail) {
                Subject = model.title,
                Body = model.content,
                BodyEncoding = Encoding.UTF8,
                IsBodyHtml = true
            };
            client.Send(message);
        }

        protected void GetMailSend()
        {
            btnOK.Enabled = false;
            int progress = 0;
            progressBar.Value = 0;
            int step = 100 / recieverGridView.Rows.Count;
            //txtSender.Text = txtSender.Text.TrimEnd( "@qq.com".ToCharArray()) + @"@qq.com";
            SmtpClient client = new SmtpClient {
                Host = "smtp.qq.com" ,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(txtSender.Text, txtPwd.Text),
                DeliveryMethod = SmtpDeliveryMethod.Network
            };
            string logSuccess = "发送成功:" + Resources.Endline;
            string logFailed = "发送失败:" + Resources.Endline;

            BackgroundWorker bgWork = new BackgroundWorker {WorkerReportsProgress = true};

            bgWork.DoWork += (sender, e) =>
            {
                SendEmailModel model = new SendEmailModel {
                    title = "",
                    content = "",
                    recieverEmail = ""
                };
                foreach (DataGridViewRow row in recieverGridView.Rows) {
                    if ( (bool)row.Cells[0].Value == true ) {
                        try {
                            Invoke( new Action<Label , string>( ( label , str ) => {
                                label.Text = str;
                            } ) , status , "正在发送给" + row.Cells[1].Value );
                            DataRow[] results1 = _detailTable.Select( $"教师姓名 = '{row.Cells[1].Value}' " );
                            DataTable t1 = _detailTable.Clone();
                            foreach ( DataRow r in results1 ) {
                                t1.ImportRow( r );
                            }

                            DataRow[] results2 = _calculateTable.Select( $"教师姓名 = '{row.Cells[1].Value}' " );
                            DataTable t2 = _calculateTable.Clone();
                            foreach ( DataRow r in results2 ) {
                                t2.ImportRow( r );
                            }

                            model = new SendEmailModel {
                                title = Convert.ToString( "课时，计算" ) ,
                                content = GenerateContent( t1 ).OuterXml + Resources.Endline + Resources.Endline+ GenerateContent( t2 ).OuterXml ,
                                recieverEmail = row.Cells[3].Value.ToString()
                            };
                            SendTo( client , model );
                            logSuccess += row.Cells[1].Value.ToString() + "\t" + row.Cells[3].Value.ToString() + Resources.Endline;
                            row.Cells[0].Value = false;
                            bgWork.ReportProgress( progress += step , model );
                        }
                        catch ( Exception ex ) {
                            logFailed += row.Cells[1].Value.ToString() + "\t" + row.Cells[3].Value.ToString() + Resources.Endline + ex.Message;
                        }
                    }
                }
            };
            bgWork.ProgressChanged += (sender, e) => {
                progressBar.Value = e.ProgressPercentage;
                txtTitle.Text = ((SendEmailModel)(e.UserState)).title;
                contentEditor.Text = ((SendEmailModel)(e.UserState)).content;
            };
            bgWork.RunWorkerCompleted += (sender, e) => {
                progressBar.Value = 100;
                btnOK.Enabled = true;
                status.Text = "完成";
                MessageBox.Show(logSuccess + logFailed);
            };
            bgWork.RunWorkerAsync();
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
                        td.Attributes.Append(createAttribute(xml, "style", STYLE_HEADER));
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
                        td.Attributes.Append(createAttribute(xml, "style", STYLE_CONTENT));
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
        }
    }
}

