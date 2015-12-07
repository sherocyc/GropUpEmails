using GropUpEmails.Properties;
using System;
using System.ComponentModel;
using System.Data;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;

namespace GropUpEmails {
    class GroupUpEmailManager {
        MainForm ui;
        EmailEngine engine;
        public int Progress{
            set {
                ui.progressBar.Value = value;
            }
        }
        public GroupUpEmailManager ( MainForm mainForm ) {
            ui = mainForm;
            engine = new EmailEngine(this); 
        }
        public void UpdateRecievers ( string xlsFilePath , Action<DataTable> callback ) {
            engine.UpdateRecievers( xlsFilePath , callback );
        }
        public void UpdateData ( string xlsFilePath , Action callback ) {
            engine.UpdateData( xlsFilePath , callback );
        }
        public void GetPreview ( int row ,Action<string , string> callback ) {
            engine.MakeEmailData( row , callback );
        }

        public void StartGroupSend (Action<string> onStatusChange , Action<string,string> onMakePreview , Action<string> onStop ) {
            Progress = 0;

            float progress = 0;
            int count = 0;

            foreach ( DataGridViewRow row in ui.recieverGridView.Rows )
                if ( (bool)row.Cells[0].Value == true )
                    count++;

            float step = 100.0f / count ;

            SmtpClient client = new SmtpClient {
                Host = ((ComboBoxItem)ui.comboBox.SelectedItem).Value,
                UseDefaultCredentials = false ,
                Credentials = new NetworkCredential( ui.txtSender.Text + ui.comboBox.SelectedItem , ui.txtPwd.Text ) ,
                DeliveryMethod = SmtpDeliveryMethod.Network
            };
            string logSuccess = "发送成功:" + Resources.Endline;
            string logFailed = "发送失败:" + Resources.Endline;

            BackgroundWorker bgWork = new BackgroundWorker { WorkerReportsProgress = true };

            bgWork.DoWork += ( sender , e ) => {
                SendEmailModel model = new SendEmailModel {
                    title = "" ,
                    content = "" ,
                    recieverEmail = ""
                };
                foreach ( DataGridViewRow row in ui.recieverGridView.Rows ) {
                    if ( (bool)row.Cells[0].Value == true ) {
                        try {
                            ui.Invoke( onStatusChange , "正在发送给" + row.Cells[1].Value );
                            engine.MakeEmailData( row.Index , ( title , content ) => {
                                model = new SendEmailModel {
                                    title = title,
                                    content = content ,
                                    recieverEmail = row.Cells[3].Value.ToString()
                                };
                            } );
                
                            engine.DoSendAction( client , model );
                            logSuccess += row.Cells[1].Value.ToString() + "\t" + row.Cells[3].Value.ToString() + Resources.Endline;
                            row.Cells[0].Value = false;
                            bgWork.ReportProgress( (int)(progress += step) , model );
                        }
                        catch ( Exception ex ) {
                            logFailed += row.Cells[1].Value.ToString() + "\t" + row.Cells[3].Value.ToString() + Resources.Endline + ex.Message;
                            break;
                        }
                    }
                }
            };
            bgWork.ProgressChanged += ( sender , e ) => {
                Progress = e.ProgressPercentage;
                onMakePreview?.Invoke( ( (SendEmailModel)( e.UserState ) ).title , ( (SendEmailModel)( e.UserState ) ).content );
            };
            bgWork.RunWorkerCompleted += ( sender , e ) => {
                Progress = 100;
                onStop?.Invoke( logSuccess + Resources.Endline + logFailed );
                onStatusChange?.Invoke( "完成" );
            };
            bgWork.RunWorkerAsync();
        }
    }
}
