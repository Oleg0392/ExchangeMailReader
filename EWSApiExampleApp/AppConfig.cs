using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;

namespace EWSApiExampleApp
{
    class AppConfig
    {
        public string configPath;
        public string parrentFolderPath;
        public int exchangeVersion;
        public string fileExtension;
        public int webCredentialsType;
        public string username;
        public string password;
        public string domain;
        public string serviceUrl;
        public string userEmailAddress;
        public int viewPageSize;

        //Config logs
        public string LogExceptionMessage1;
        public string LogFileWritingComplete;

        //Regex XML patterns
        private string mainPattern = @"<[A-z]+\d*>\S*.+:?</[A-z]+\d*>";
        private string namePattern = @"<[A-z]+\d*>";
        private string valuePattern = @">\S*.+:?<";

        //Config dictionary
        private string[] settingNames = new string[20];
        private string[] settingValues = new string[20];

        public AppConfig(string filePath)
        {
            configPath = filePath;
            StreamReader streamReader;
            int settingIndex = 0;

            Regex regexMain = new Regex(mainPattern);
            Regex regexSettingName = new Regex(namePattern);
            Regex regexSettingValue = new Regex(valuePattern);

            if (!(String.IsNullOrEmpty(configPath)))
            {
                try
                {
                    streamReader = new StreamReader(filePath);
                    string configText = streamReader.ReadToEnd();

                    MatchCollection matchCollection = regexMain.Matches(configText);
                    foreach (Match match in matchCollection)
                    {
                        string currentItem = match.Value;
                        string ItemName = regexSettingName.Match(currentItem).Value;
                        string ItemValue = regexSettingValue.Match(currentItem).Value;

                        settingNames[settingIndex] = ItemName.Substring(1, ItemName.Length - 2);
                        settingValues[settingIndex] = ItemValue.Substring(1, ItemValue.Length - 2);

                        settingIndex++;
                    }
                } catch (Exception e) { Console.WriteLine($"{e.Message} Source: {e.Source}"); }

                //Parsing config fields with reflection tools         
                Type ConfigType = typeof(AppConfig);
                FieldInfo[] fieldInfos = ConfigType.GetFields(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly | BindingFlags.NonPublic);
                foreach (FieldInfo currentField in fieldInfos)
                {
                    int needIndex = getSettingIndex(currentField.Name);
                    if (needIndex == -1) continue;                                //Если данное поле не я вляется настройкой или оно private
                    if (currentField.FieldType == needIndex.GetType())
                    {
                        currentField.SetValue(this, Convert.ToInt32(settingValues[needIndex]));
                    }
                    else currentField.SetValue(this, settingValues[needIndex]);
                }

            }
            else
            {
                throw new ArgumentException(filePath,"Invalid path to the config file!");
            }
 
        }


        private int getSettingIndex(string settingName)
        {
            int valueIndex = -1;

            for (int i = 0; i < settingNames.Length; i++)
            {
                if (settingName.Equals(settingNames[i]))
                {
                    valueIndex = i;
                }

            }

            //if (valueIndex == -1) throw new ArgumentException(settingName, "Setting name is incorrect or absent in to the settings array.");

            return valueIndex;
        }



        //Parsing config settings without Reflection
        /*parrentFolderPath = settingValues[getSettingValue("parrentFolderPath")];
          exchangeVersion = Convert.ToInt32(settingValues[getSettingValue("exchangeVersion")]);
          fileExtension = settingValues[getSettingValue("fileExtension")];
          webCredentialsType = Convert.ToInt32(settingValues[getSettingValue("webCredentialsType")]);
          if (webCredentialsType == 2)
          {
              username = settingValues[getSettingValue("username")];
              password = settingValues[getSettingValue("password")];
              domain = settingValues[getSettingValue("domain")];
          }
          serviceUrl = settingValues[getSettingValue("serviceUrl")];
          userEmailAddress = settingValues[getSettingValue("userEmailAddress")];
          viewPageSize = Convert.ToInt32(settingValues[getSettingValue("viewPageSize")]);

          LogExceptionMessage1 = settingValues[getSettingValue("LogExceptionMessage1")];
          //LogFileWritingComplete = settingValues[getSettingValue("LogFileWritingComplete")];*/
    }
}
