using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using GropUpEmails.Properties;

namespace GropUpEmails
{
    public partial class GroupUpEmails : Form
    {
        public GroupUpEmails()
        {
            InitializeComponent();
            btnReciever.Click += (sender, e) =>{
                GetRecieverFilePath();
            };
            btnData.Click += (sender, e) => {
                GetDataFilePath();
            };
            btnCancel.Click += (sender, e) => {
                ClearAllText();
            };
            btnOK.Click += (sender, e) => {
                GetMailSend(GetAdressOfRecievers(txtRecieverFile.Text));
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
        private void GetContent() {

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
        private void ClearAllText()
        {
            FieldInfo[] infos = GetType().GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.GetField | BindingFlags.Instance);
            for (int i = 0; i < infos.Length; i++)
            {
                if (infos[i].FieldType == typeof(TextBox))
                {
                    ((TextBox)infos[i].GetValue(this)).Text = "";
                }
            }
        }

        protected void GetMailSend(ArrayList alEmail)
        {
            try
            {
                txtSender.Text = txtSender.Text.TrimEnd("@qq.com".ToCharArray()) + @"@qq.com";
                SmtpClient client = new SmtpClient
                {
                    Host = "smtp.qq.com",
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(txtSender.Text, txtPwd.Text),
                    DeliveryMethod = SmtpDeliveryMethod.Network
                };
                var array = alEmail;
                foreach(string email in array )
                {
                    MailMessage message = new MailMessage(txtSender.Text, email)
                    {
                        Subject = txtTitle.Text,
                        Body = txtContent.Text,
                        BodyEncoding = Encoding.UTF8,
                        IsBodyHtml = true
                    };
                    client.Send(message);
                }
                MessageBox.Show("发送邮件成功！");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        protected  ArrayList GetAdressOfRecievers(string xlsFilePath)
        {
            string strConn = $"Provider=Microsoft.Ace.OleDb.12.0;Data Source={xlsFilePath};Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";
            string strComm = "SELECT * FROM [Sheet1$]";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbDataAdapter myAdp = new OleDbDataAdapter(strComm, strConn);
            DataSet ds = new DataSet();
            myAdp.Fill(ds);
            myConn.Close();
            
            DataTable table = ds.Tables[0];
                
            ArrayList email = new ArrayList();
            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count;i++ )
                {
                    String strEmail = Convert.ToString(row[i]);
                    if(strEmail .Contains('@')&&strEmail.Contains ('.'))
                    {
                        email.Add(strEmail);
                    }
                }
            }
            return email;
        }
    }
}

