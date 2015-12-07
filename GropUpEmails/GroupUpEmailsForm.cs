using System.Windows.Forms;
using GropUpEmails.Properties;
using Application = System.Windows.Forms.Application;

namespace GropUpEmails
{
    public class ComboxItem {
        public string Text {
            get; set;
        }
        public string Value {
            get; set;
        }

        public ComboxItem ( string _Text , string _Value ) {
            Text = _Text;
            Value = _Value;
        }

        public override string ToString () {
            return Text;
        }
    }


    public partial class GroupUpEmailsForm : Form
    {
        private GroupUpEmailManager manager;

        public GroupUpEmailsForm()
        {
            InitializeComponent();
            manager = new GroupUpEmailManager( this );

            txtSender.Text = UserPreference.Instance.Data.UserEmail;
            txtPwd.Text = UserPreference.Instance.Data.UserEmailPassword;

            btnData.Enabled = false;
            regenarateBtn.Enabled = false;

            ComboxItem[] items = {
                new ComboxItem( "@qq.com" , "smtp.qq.com" ),
                new ComboxItem( "@163.com" , "smtp.163.com" ),
            };
            comboBox.Items.AddRange( items );
            comboBox.SelectedIndex = UserPreference.Instance.Data.SmtpIndex;
            btnReciever.Click += (sender, e) =>{
                Utils.OpenXlsFile(filename => {
                    txtRecieverFile.Text = filename;
                    manager.UpdateRecievers( filename , table => {
                        recieverGridView.Columns.Clear();
                        recieverGridView.Columns.Add( new DataGridViewCheckBoxColumn {
                            HeaderText = "选择" ,
                            Name = "选择" ,
                            ReadOnly = false
                        } );
                        recieverGridView.DataSource = table;
                        recieverGridView.CurrentCell = recieverGridView.Rows[0].Cells[0];
                        recieverGridView.Columns[0].Width = 30;
                        foreach ( DataGridViewRow row in recieverGridView.Rows ) {
                            row.Cells[0].Value = true;
                        }
                        btnData.Enabled = true;
                    } );
                });
            };

            btnData.Click += (sender, e) => {
                Utils.OpenXlsFile( filename => {
                    txtDataFile.Text = filename;
                    manager.UpdateData( filename , () => {
                        manager.GetPreview( recieverGridView.SelectedRows[0].Index,( title , content ) => {
                            txtTitle.Text = title;
                            contentEditor.Text = content;
                        } );
                    } );
                    regenarateBtn.Enabled = true;
                } );
            };

            regenarateBtn.Click += (sender, e) => {
                manager.GetPreview( recieverGridView.SelectedRows[0].Index , ( title , content ) => {
                    txtTitle.Text = title;
                    contentEditor.Text = content;
                } );
            };

            btnOK.Click += ( sender , e ) => {
                btnOK.Enabled = false;
                manager.StartGroupSend( state => {
                    status.Text = state;
                } , ( title,content )=> {
                    txtTitle.Text = title;
                    contentEditor.Text = content ;
                } , result => {
                    btnOK.Enabled = true;
                    MessageBox.Show( result );
                } );
            };

            FormClosing += (sender, e) => {
                if (MessageBox.Show(Resources.ComfirmQuit, Resources.Quit, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    e.Cancel = true;
                    WindowState = FormWindowState.Minimized;
                    ShowInTaskbar = false;
                    UserPreference.Instance.Data.UserEmail = txtSender.Text;
                    UserPreference.Instance.Data.UserEmailPassword = txtPwd.Text;
                    UserPreference.Instance.Data.RecieverDataPath = txtRecieverFile.Text ;
                    UserPreference.Instance.Data.SmtpIndex = comboBox.SelectedIndex;
                    UserPreference.Instance.Save();
                    Application.Exit();
                }
                else {
                    e.Cancel = true;
                }
            };
        }
    }
}

