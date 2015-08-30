using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Collections;

namespace GropUpEmails
{
    public partial class GroupUpEmails : Form
    {
        public GroupUpEmails()
        {
            InitializeComponent();
        }

        private void GroupUpEmails_Load(object sender, EventArgs e)
        {
        }
        private void btnReciever_Click(object sender, EventArgs e)
        {
            GetRecieverFilePath();
        }
        private void GetRecieverFilePath()
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            // Insert code to read the stream here.
                        }
                        this.txtRecieverFile.Text = openFileDialog1.FileName;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("文件打开错误： " + ex.Message);
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            GetAllTextClear();
        }
        private void GetAllTextClear()
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

        private void btnOK_Click(object sender, EventArgs e)
        {
            GetMailSend(GetAdressOfRecievers(txtRecieverFile.Text ));
        }
        protected void GetMailSend(ArrayList alEmail)
        {
            try
            {
                txtSender.Text = txtSender.Text.TrimEnd("@163.com".ToCharArray()) + "@163.com";
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = "smtp.163.com";
                client.UseDefaultCredentials = false;
                client.Credentials =new System.Net.NetworkCredential(txtSender.Text, txtPwd.Text);
                client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                ArrayList array =new ArrayList();
                array = alEmail;
                foreach(string email in array )
                {
                //System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(txtSender.Text, "HeXiaoMing408@163.com");
                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(txtSender.Text, email);
                message.Subject = txtTitle.Text;
                message.Body = txtContent.Text;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.IsBodyHtml = true;
                client.Send(message);
                }
                MessageBox.Show("发送邮件成功！");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        protected  ArrayList GetAdressOfRecievers(string XlsFilePath)
        {
            try
            {
                string strConn = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=" + XlsFilePath + ";Extended Properties=Excel 8.0";
                string strComm = "SELECT * FROM [Sheet1$]";
                OleDbConnection myConn = new OleDbConnection(strConn);
                myConn.Open();
                OleDbDataAdapter myAdp = new OleDbDataAdapter(strComm, strConn);
                DataSet ds = new DataSet();
                myAdp.Fill(ds);
                myConn.Close();
                DataTable table = ds.Tables[0];
                
                ArrayList Email = new ArrayList();
                foreach (DataRow row in table.Rows)
                {
                    for (int i = 0; i < table.Columns.Count;i++ )
                    {
                        string strEmail = string.Empty;
                        //strEmail = Convert.ToString(row[0]);
                        strEmail = Convert.ToString(row[i]);
                        if(strEmail .Contains('@')&&strEmail.Contains ('.'))
                        {
                        Email.Add(strEmail);
                        }
                    }
                }
                return Email;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw e;
            }
            finally
            {
                
            }
        }

        private void GroupUpEmails_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("确定需要退出窗口？", "确认退出...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
            }
            else
            {
                e.Cancel = true;
            }
        }

    }
}

