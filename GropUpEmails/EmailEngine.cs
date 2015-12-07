using GropUpEmails.Properties;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace GropUpEmails {
    class EmailEngine {

        private const string STYLE_HEADER =
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

        private DataTable _recieverTable;

        private DataTable _detailTable;
        private DataTable _calculateTable;

        GroupUpEmailManager manager;

        public EmailEngine ( GroupUpEmailManager manager ) {
            this.manager = manager;
        }

        public void DoSendAction ( SmtpClient client , SendEmailModel model ) {
            MailMessage message = new MailMessage( ((NetworkCredential)client.Credentials).UserName, model.recieverEmail ) {
                Subject = model.title ,
                Body = model.content ,
                BodyEncoding = Encoding.UTF8 ,
                IsBodyHtml = true
            };
            client.Send( message );
        }

        public void UpdateRecievers ( string xlsFilePath , Action<DataTable> callback ) {
            try {
                manager.Progress = 0;
                string strConn = $"Provider=Microsoft.Ace.OleDb.12.0;Data Source={xlsFilePath};Extended Properties='Excel 12.0; HDR=Yes; IMEX=1;'";
                string strComm = "SELECT 教师姓名,教师证件号,邮箱号 FROM [Sheet1$]";
                OleDbConnection myConn = new OleDbConnection( strConn );
                OleDbDataAdapter myAdp = new OleDbDataAdapter( strComm , strConn );
                DataSet ds = new DataSet();
                myAdp.Fill( ds );
                myConn.Close();
                manager.Progress = 50;

                _recieverTable = ds.Tables[0];

                if ( callback != null )
                    callback( _recieverTable );

                manager.Progress = 100;
            }
            catch ( Exception e ) {
                MessageBox.Show( e.Message );
            }
        }

        public void UpdateData ( string xlsFilePath , Action callback ) {
            BackgroundWorker bgWork = new BackgroundWorker();
            bgWork.DoWork += ( sender , e ) => {
                try {
                    string strConn = $"Provider=Microsoft.Ace.OleDb.12.0;Data Source={xlsFilePath};Extended Properties='Excel 12.0; HDR=Yes;IMEX=1;'";
                    string strComm = "SELECT * FROM [英语$] ";
                    OleDbConnection myConn = new OleDbConnection( strConn );
                    myConn.Open();

                    OleDbDataAdapter myAdp = new OleDbDataAdapter( strComm , strConn );
                    DataSet ds = new DataSet();
                    myAdp.Fill( ds );
                    _detailTable = ds.Tables[0];

                    strComm = "SELECT * FROM [计算列$] ";
                    myAdp = new OleDbDataAdapter( strComm , strConn );
                    ds = new DataSet();
                    myAdp.Fill( ds );
                    _calculateTable = ds.Tables[0];

                    myConn.Close();
                }
                catch ( Exception ex ) {
                    MessageBox.Show( ex.Message );
                }
            };
            bgWork.RunWorkerCompleted += ( sender , e ) => {
                //callback?.Invoke();
                if ( callback != null )
                    callback();
            };
            bgWork.RunWorkerAsync();
        }

        public void MakeEmailData (int rowIndex , Action<string,string> callback) {
            try {
                string name = Convert.ToString( _recieverTable.Rows[rowIndex][0] );

                DataRow[] results1 = _detailTable.Select( $"教师姓名 = '{name}' " );
                DataTable t1 = _detailTable.Clone();
                foreach ( DataRow r in results1 ) {
                    t1.ImportRow( r );
                }

                DataRow[] results2 = _calculateTable.Select( $"教师姓名 = '{name}' " );
                DataTable t2 = _calculateTable.Clone();
                foreach ( DataRow r in results2 ) {
                    t2.ImportRow( r );
                }
                string title = Convert.ToString( name + "  课时，计算" );
                string content = GenerateContent( t1 ).OuterXml + Resources.Endline + Resources.Endline + GenerateContent( t2 ).OuterXml;
                callback?.Invoke( title , content);
            }
            catch ( Exception e ) {
                MessageBox.Show( e.Message );
            }
        }

        private XmlAttribute createAttribute ( XmlDocument doc , string name , string value ) {
            XmlAttribute attr = doc.CreateAttribute( name );
            attr.Value = value;
            return attr;
        }
        private XmlDocument GenerateContent ( DataTable data ) {
            try {
                XmlDocument xml = new XmlDocument();
                XmlNode div = xml.CreateElement( "div" );
                XmlNode table = xml.CreateElement( "table" );
                table.Attributes.Append( createAttribute( xml , "style" , "WIDTH: 919pt; BORDER-COLLAPSE: collapse" ) );
                table.Attributes.Append( createAttribute( xml , "cellspacing" , "0" ) );
                table.Attributes.Append( createAttribute( xml , "cellpadding" , "0" ) );
                table.Attributes.Append( createAttribute( xml , "width" , "1225" ) );
                table.Attributes.Append( createAttribute( xml , "border" , "0" ) );
                XmlNode colgroup = xml.CreateElement( "colgroup" );
                foreach ( var column in data.Columns ) {
                    XmlNode col = xml.CreateElement( "col" );
                    col.Attributes.Append( createAttribute( xml , "style" , "WIDTH: 92pt; mso-width-source: userset; mso-width-alt: 4380" ) );
                    col.Attributes.Append( createAttribute( xml , "width" , "123" ) );
                    colgroup.AppendChild( col );
                }
                XmlNode tbody = xml.CreateElement( "tbody" );

                {
                    XmlNode tr = xml.CreateElement( "tr" );
                    tr.Attributes.Append( createAttribute( xml , "style" , "HEIGHT: 23.25pt; mso-height-source: userset" ) );
                    tr.Attributes.Append( createAttribute( xml , "height" , "30" ) );
                    tbody.AppendChild( tr );
                    foreach ( var column in data.Columns ) {
                        XmlNode td = xml.CreateElement( "td" );
                        td.Attributes.Append( createAttribute( xml , "class" , "xl23212" ) );
                        td.Attributes.Append( createAttribute( xml , "style" , STYLE_HEADER ) );
                        td.Attributes.Append( createAttribute( xml , "height" , "30" ) );
                        td.Attributes.Append( createAttribute( xml , "width" , "123" ) );
                        XmlNode strong = xml.CreateElement( "strong" );
                        XmlNode font = xml.CreateElement( "font" );
                        font.Attributes.Append( createAttribute( xml , "size" , "2" ) );
                        font.Attributes.Append( createAttribute( xml , "face" , "宋体" ) );
                        font.InnerText = column.ToString();
                        strong.AppendChild( font );
                        td.AppendChild( strong );
                        tbody.AppendChild( td );
                    }
                }

                foreach ( DataRow row in data.Rows ) {
                    XmlNode tr = xml.CreateElement( "tr" );
                    tr.Attributes.Append( createAttribute( xml , "style" , "HEIGHT: 12pt" ) );
                    tr.Attributes.Append( createAttribute( xml , "height" , "16" ) );
                    tbody.AppendChild( tr );
                    for ( int i = 0 ; i < data.Columns.Count ; i++ ) {
                        XmlNode td = xml.CreateElement( "td" );
                        td.Attributes.Append( createAttribute( xml , "class" , "xl23212" ) );
                        td.Attributes.Append( createAttribute( xml , "style" , STYLE_CONTENT ) );
                        td.Attributes.Append( createAttribute( xml , "height" , "16" ) );
                        td.Attributes.Append( createAttribute( xml , "width" , "123" ) );
                        XmlNode font = xml.CreateElement( "font" );
                        font.Attributes.Append( createAttribute( xml , "size" , "2" ) );
                        font.Attributes.Append( createAttribute( xml , "face" , "宋体" ) );
                        font.InnerText = row[i].ToString();
                        td.AppendChild( font );
                        tbody.AppendChild( td );
                    }
                }
                table.AppendChild( tbody );
                table.AppendChild( colgroup );
                div.AppendChild( table );
                xml.AppendChild( div );
                return xml;
            }
            catch ( Exception e ) {
                MessageBox.Show( e.Message );
                throw;
            }
        }

    }
}
