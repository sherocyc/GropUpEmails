using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Drawing ;

namespace GropUpEmails
{
    public class Tray:Form
    {
        public bool FormGroupUpEmailExist = true;
        private Icon groupUpEmailIcon = new Icon("GroupUpEmail.ico");
        private NotifyIcon TrayIcon;
        private ContextMenu notifyiconMenu;
        GropUpEmails.GroupUpEmails newGroupUpEmail = new GroupUpEmails();
        public Tray()
        {
            InitializeNotifyico();
            InitializeComponent();
            newGroupUpEmail.Show();
        }
        private void InitializeNotifyico()
        {
            TrayIcon = new NotifyIcon();
            TrayIcon.Icon = groupUpEmailIcon;
            TrayIcon.Text = "群起邮件"+"@_Hirisw_@"+"作者：何晓明200910";
            TrayIcon.Visible = true;
            TrayIcon.Click +=new System.EventHandler(this.click);

            MenuItem[] miItems=new MenuItem [3];
  
            miItems[0]=new MenuItem();
            miItems[0].Text="群起邮件";
            miItems[0].Click +=new System.EventHandler (this.showmessage);

            miItems[1]=new MenuItem ("-");
            
            miItems[2]=new MenuItem ();
            miItems[2].Text="退出程序";
            miItems[2].Click +=new System.EventHandler (this.ExitSelect );

            notifyiconMenu =new ContextMenu (miItems );
            TrayIcon.ContextMenu=notifyiconMenu ;
        }
        public void click(object sender, System.EventArgs e)
        {
            newGroupUpEmail.Show();
            newGroupUpEmail.WindowState = FormWindowState.Normal;
            newGroupUpEmail.ShowInTaskbar = true;
        }
        public void showmessage(object sender, EventArgs e)
        {
            newGroupUpEmail.Show();
        }
        public void ExitSelect(object sender, EventArgs e)
        {
            TrayIcon.Visible = false;
            this.Close();
        }
        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(320, 56);
            this.ControlBox = false;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.WindowState = FormWindowState.Minimized;

            this.Name = "Tray";
            this.ShowInTaskbar = false;
            this.Text = "Tray For GroupUpEmail";
            this.ResumeLayout(false);
        }
    }
}
