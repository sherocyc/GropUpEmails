using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace GropUpEmails
{
    partial class GroupUpEmails
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private IContainer components = null;

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
            this.label1 = new System.Windows.Forms.Label();
            this.txtDataFile = new System.Windows.Forms.TextBox();
            this.btnData = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.recieverGridView = new System.Windows.Forms.DataGridView();
            this.regenarateBtn = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.recieverGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // lblSender
            // 
            this.lblSender.AutoSize = true;
            this.lblSender.Location = new System.Drawing.Point(16, 14);
            this.lblSender.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblSender.Name = "lblSender";
            this.lblSender.Size = new System.Drawing.Size(97, 15);
            this.lblSender.TabIndex = 0;
            this.lblSender.Text = "发件人邮箱：";
            // 
            // txtSender
            // 
            this.txtSender.Location = new System.Drawing.Point(231, 10);
            this.txtSender.Margin = new System.Windows.Forms.Padding(4);
            this.txtSender.Name = "txtSender";
            this.txtSender.Size = new System.Drawing.Size(425, 25);
            this.txtSender.TabIndex = 2;
            // 
            // lblReciever
            // 
            this.lblReciever.AutoSize = true;
            this.lblReciever.Location = new System.Drawing.Point(16, 75);
            this.lblReciever.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblReciever.Name = "lblReciever";
            this.lblReciever.Size = new System.Drawing.Size(175, 15);
            this.lblReciever.TabIndex = 3;
            this.lblReciever.Text = "收件人邮箱（*.xlsx）：";
            // 
            // openFileDialogRecievers
            // 
            this.openFileDialogRecievers.FileName = "openFileDialogRecievers";
            this.openFileDialogRecievers.ShowHelp = true;
            this.openFileDialogRecievers.SupportMultiDottedExtensions = true;
            // 
            // btnReciever
            // 
            this.btnReciever.Location = new System.Drawing.Point(664, 69);
            this.btnReciever.Margin = new System.Windows.Forms.Padding(4);
            this.btnReciever.Name = "btnReciever";
            this.btnReciever.Size = new System.Drawing.Size(160, 25);
            this.btnReciever.TabIndex = 4;
            this.btnReciever.Text = "打开收件人文件...";
            this.btnReciever.UseVisualStyleBackColor = true;
            // 
            // txtRecieverFile
            // 
            this.txtRecieverFile.Location = new System.Drawing.Point(231, 69);
            this.txtRecieverFile.Margin = new System.Windows.Forms.Padding(4);
            this.txtRecieverFile.Name = "txtRecieverFile";
            this.txtRecieverFile.ReadOnly = true;
            this.txtRecieverFile.Size = new System.Drawing.Size(425, 25);
            this.txtRecieverFile.TabIndex = 5;
            // 
            // lblContent
            // 
            this.lblContent.AutoSize = true;
            this.lblContent.Location = new System.Drawing.Point(16, 166);
            this.lblContent.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblContent.Name = "lblContent";
            this.lblContent.Size = new System.Drawing.Size(82, 15);
            this.lblContent.TabIndex = 6;
            this.lblContent.Text = "内容预览：";
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Location = new System.Drawing.Point(16, 138);
            this.lblTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(82, 15);
            this.lblTitle.TabIndex = 8;
            this.lblTitle.Text = "标题预览：";
            // 
            // txtContent
            // 
            this.txtContent.Location = new System.Drawing.Point(111, 168);
            this.txtContent.Margin = new System.Windows.Forms.Padding(4);
            this.txtContent.Multiline = true;
            this.txtContent.Name = "txtContent";
            this.txtContent.ReadOnly = true;
            this.txtContent.Size = new System.Drawing.Size(714, 395);
            this.txtContent.TabIndex = 9;
            // 
            // txtTitle
            // 
            this.txtTitle.Location = new System.Drawing.Point(111, 135);
            this.txtTitle.Margin = new System.Windows.Forms.Padding(4);
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.ReadOnly = true;
            this.txtTitle.Size = new System.Drawing.Size(545, 25);
            this.txtTitle.TabIndex = 10;
            // 
            // lblPwd
            // 
            this.lblPwd.AutoSize = true;
            this.lblPwd.Location = new System.Drawing.Point(16, 44);
            this.lblPwd.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblPwd.Name = "lblPwd";
            this.lblPwd.Size = new System.Drawing.Size(127, 15);
            this.lblPwd.TabIndex = 11;
            this.lblPwd.Text = "发件人邮箱密码：";
            // 
            // txtPwd
            // 
            this.txtPwd.Location = new System.Drawing.Point(231, 40);
            this.txtPwd.Margin = new System.Windows.Forms.Padding(4);
            this.txtPwd.Name = "txtPwd";
            this.txtPwd.PasswordChar = '*';
            this.txtPwd.Size = new System.Drawing.Size(425, 25);
            this.txtPwd.TabIndex = 12;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(721, 571);
            this.btnOK.Margin = new System.Windows.Forms.Padding(4);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(100, 29);
            this.btnOK.TabIndex = 13;
            this.btnOK.Text = "确定发送";
            this.btnOK.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(661, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 15);
            this.label1.TabIndex = 15;
            this.label1.Text = "@juren.com";
            // 
            // txtDataFile
            // 
            this.txtDataFile.Location = new System.Drawing.Point(231, 102);
            this.txtDataFile.Margin = new System.Windows.Forms.Padding(4);
            this.txtDataFile.Name = "txtDataFile";
            this.txtDataFile.ReadOnly = true;
            this.txtDataFile.Size = new System.Drawing.Size(425, 25);
            this.txtDataFile.TabIndex = 18;
            // 
            // btnData
            // 
            this.btnData.Location = new System.Drawing.Point(664, 102);
            this.btnData.Margin = new System.Windows.Forms.Padding(4);
            this.btnData.Name = "btnData";
            this.btnData.Size = new System.Drawing.Size(160, 25);
            this.btnData.TabIndex = 17;
            this.btnData.Text = "打开数据文件...";
            this.btnData.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 108);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(160, 15);
            this.label2.TabIndex = 16;
            this.label2.Text = "内容数据（*.xlsx）：";
            // 
            // recieverGridView
            // 
            this.recieverGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.recieverGridView.Location = new System.Drawing.Point(843, 14);
            this.recieverGridView.MultiSelect = false;
            this.recieverGridView.Name = "recieverGridView";
            this.recieverGridView.RowTemplate.Height = 27;
            this.recieverGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.recieverGridView.Size = new System.Drawing.Size(313, 587);
            this.recieverGridView.TabIndex = 19;
            // 
            // regenarateBtn
            // 
            this.regenarateBtn.Location = new System.Drawing.Point(664, 135);
            this.regenarateBtn.Margin = new System.Windows.Forms.Padding(4);
            this.regenarateBtn.Name = "regenarateBtn";
            this.regenarateBtn.Size = new System.Drawing.Size(160, 25);
            this.regenarateBtn.TabIndex = 20;
            this.regenarateBtn.Text = "重新生成预览";
            this.regenarateBtn.UseVisualStyleBackColor = true;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(111, 575);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(590, 23);
            this.progressBar.TabIndex = 21;
            // 
            // GroupUpEmails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1168, 613);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.regenarateBtn);
            this.Controls.Add(this.recieverGridView);
            this.Controls.Add(this.txtDataFile);
            this.Controls.Add(this.btnData);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
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
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "GroupUpEmails";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "吕陶丽群发邮件";
            ((System.ComponentModel.ISupportInitialize)(this.recieverGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Label lblSender;
        private TextBox txtSender;
        private Label lblReciever;
        private OpenFileDialog openFileDialogRecievers;
        private Button btnReciever;
        private TextBox txtRecieverFile;
        private Label lblContent;
        private Label lblTitle;
        private TextBox txtContent;
        private TextBox txtTitle;
        private Label lblPwd;
        private TextBox txtPwd;
        private Button btnOK;
        private Label label1;
        private TextBox txtDataFile;
        private Button btnData;
        private Label label2;
        private DataGridView recieverGridView;
        private Button regenarateBtn;
        private ProgressBar progressBar;
    }
}

