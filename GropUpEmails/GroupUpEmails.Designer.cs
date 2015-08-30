namespace GropUpEmails
{
    partial class GroupUpEmails
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GroupUpEmails));
            this.lblSender = new System.Windows.Forms.Label();
            this.txtSender = new System.Windows.Forms.TextBox();
            this.lblReciever = new System.Windows.Forms.Label();
            this.openFileDialogRecievers = new System.Windows.Forms.OpenFileDialog();
            this.btnReciever = new System.Windows.Forms.Button();
            this.txtRecieverFile = new System.Windows.Forms.TextBox();
            this.lblContent = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.txtContent = new System.Windows.Forms.TextBox();
            this.txtTitle = new System.Windows.Forms.TextBox();
            this.lblPwd = new System.Windows.Forms.Label();
            this.txtPwd = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblMail = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblSender
            // 
            this.lblSender.AutoSize = true;
            this.lblSender.Location = new System.Drawing.Point(12, 11);
            this.lblSender.Name = "lblSender";
            this.lblSender.Size = new System.Drawing.Size(155, 12);
            this.lblSender.TabIndex = 0;
            this.lblSender.Text = "发件人邮箱（限163邮箱）：";
            // 
            // txtSender
            // 
            this.txtSender.Location = new System.Drawing.Point(173, 8);
            this.txtSender.Name = "txtSender";
            this.txtSender.Size = new System.Drawing.Size(211, 21);
            this.txtSender.TabIndex = 2;
            // 
            // lblReciever
            // 
            this.lblReciever.AutoSize = true;
            this.lblReciever.Location = new System.Drawing.Point(12, 60);
            this.lblReciever.Name = "lblReciever";
            this.lblReciever.Size = new System.Drawing.Size(155, 12);
            this.lblReciever.TabIndex = 3;
            this.lblReciever.Text = "收件人文件（限Xls文件）：";
            // 
            // openFileDialogRecievers
            // 
            this.openFileDialogRecievers.FileName = "openFileDialogRecievers";
            this.openFileDialogRecievers.ShowHelp = true;
            this.openFileDialogRecievers.SupportMultiDottedExtensions = true;
            // 
            // btnReciever
            // 
            this.btnReciever.Location = new System.Drawing.Point(390, 55);
            this.btnReciever.Name = "btnReciever";
            this.btnReciever.Size = new System.Drawing.Size(120, 23);
            this.btnReciever.TabIndex = 4;
            this.btnReciever.Text = "打开收件人文件...";
            this.btnReciever.UseVisualStyleBackColor = true;
            this.btnReciever.Click += new System.EventHandler(this.btnReciever_Click);
            // 
            // txtRecieverFile
            // 
            this.txtRecieverFile.Location = new System.Drawing.Point(173, 55);
            this.txtRecieverFile.Name = "txtRecieverFile";
            this.txtRecieverFile.Size = new System.Drawing.Size(211, 21);
            this.txtRecieverFile.TabIndex = 5;
            // 
            // lblContent
            // 
            this.lblContent.AutoSize = true;
            this.lblContent.Location = new System.Drawing.Point(12, 104);
            this.lblContent.Name = "lblContent";
            this.lblContent.Size = new System.Drawing.Size(65, 12);
            this.lblContent.TabIndex = 6;
            this.lblContent.Text = "邮件内容：";
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Location = new System.Drawing.Point(12, 82);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(65, 12);
            this.lblTitle.TabIndex = 8;
            this.lblTitle.Text = "邮件标题：";
            // 
            // txtContent
            // 
            this.txtContent.Location = new System.Drawing.Point(83, 101);
            this.txtContent.Multiline = true;
            this.txtContent.Name = "txtContent";
            this.txtContent.Size = new System.Drawing.Size(427, 136);
            this.txtContent.TabIndex = 9;
            // 
            // txtTitle
            // 
            this.txtTitle.Location = new System.Drawing.Point(83, 79);
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.Size = new System.Drawing.Size(301, 21);
            this.txtTitle.TabIndex = 10;
            // 
            // lblPwd
            // 
            this.lblPwd.AutoSize = true;
            this.lblPwd.Location = new System.Drawing.Point(12, 35);
            this.lblPwd.Name = "lblPwd";
            this.lblPwd.Size = new System.Drawing.Size(101, 12);
            this.lblPwd.TabIndex = 11;
            this.lblPwd.Text = "发件人邮箱密码：";
            // 
            // txtPwd
            // 
            this.txtPwd.Location = new System.Drawing.Point(173, 32);
            this.txtPwd.Name = "txtPwd";
            this.txtPwd.PasswordChar = '*';
            this.txtPwd.Size = new System.Drawing.Size(211, 21);
            this.txtPwd.TabIndex = 12;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(343, 243);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 13;
            this.btnOK.Text = "确定发送";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(424, 243);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 14;
            this.btnCancel.Text = "取消重填";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblMail
            // 
            this.lblMail.AutoSize = true;
            this.lblMail.Location = new System.Drawing.Point(388, 11);
            this.lblMail.Name = "lblMail";
            this.lblMail.Size = new System.Drawing.Size(53, 12);
            this.lblMail.TabIndex = 15;
            this.lblMail.Text = "@163.com";
            // 
            // GroupUpEmails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(539, 271);
            this.Controls.Add(this.lblMail);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtPwd);
            this.Controls.Add(this.lblPwd);
            this.Controls.Add(this.txtTitle);
            this.Controls.Add(this.txtContent);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.lblContent);
            this.Controls.Add(this.txtRecieverFile);
            this.Controls.Add(this.btnReciever);
            this.Controls.Add(this.lblReciever);
            this.Controls.Add(this.txtSender);
            this.Controls.Add(this.lblSender);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "GroupUpEmails";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "群起邮件纪念版";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.GroupUpEmails_FormClosing);
            this.Load += new System.EventHandler(this.GroupUpEmails_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblSender;
        private System.Windows.Forms.TextBox txtSender;
        private System.Windows.Forms.Label lblReciever;
        private System.Windows.Forms.OpenFileDialog openFileDialogRecievers;
        private System.Windows.Forms.Button btnReciever;
        private System.Windows.Forms.TextBox txtRecieverFile;
        private System.Windows.Forms.Label lblContent;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.TextBox txtContent;
        private System.Windows.Forms.TextBox txtTitle;
        private System.Windows.Forms.Label lblPwd;
        private System.Windows.Forms.TextBox txtPwd;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblMail;
    }
}

