
using System.IO;

namespace GropUpEmails {
    public class UserPreferenceData {
        public string UserEmail="";
        public string UserEmailPassword="";
        public string RecieverDataPath = "";
        public int SmtpIndex = 0;
    }
    class UserPreference
    {
        private UserPreference()
        {
            Load();
        }

        private static UserPreference _instance;
        public static UserPreference Instance => _instance ?? (_instance = new UserPreference());

        private readonly string _dataFileName = Directory.GetCurrentDirectory() + "\\user.dat";//存档文件的名称,自己定//
        private readonly XmlSaver _xmlSaver = new XmlSaver();

        public UserPreferenceData Data;

        public void Save() {
            string dataString = _xmlSaver.SerializeObject(Data ?? (Data = new UserPreferenceData()), typeof(UserPreferenceData));
            _xmlSaver.CreateXML(_dataFileName, dataString);
        }

        private void Load() {
            if (File.Exists(_dataFileName)) {
                string dataString = _xmlSaver.LoadXML(_dataFileName);
                Data = _xmlSaver.DeserializeObject(dataString, typeof(UserPreferenceData)) as UserPreferenceData;
            }
            else {
                 Save();
            }
        }
    }
}
