using GropUpEmails.Properties;
using System;
using System.Windows.Forms;

namespace GropUpEmails {
    public abstract class Utils {
        public static void OpenXlsFile ( Action<string> listener) {
            OpenFileDialog openFileDialog = new OpenFileDialog {
                InitialDirectory = Environment.GetFolderPath( Environment.SpecialFolder.DesktopDirectory ) ,
                Filter = Resources.ExcelFilter ,
                FilterIndex = 1 ,
                RestoreDirectory = true
            };

            if ( openFileDialog.ShowDialog() == DialogResult.OK ){
                listener( openFileDialog.FileName );
            }
        }
    }
}
