using csharp.rpa.challenge.selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rpa.challenge.Constants
{
    class ChallengeConstants
    {
        public static string URL_CHALLENGE = GetValue("url", "mainUrlChallenge");
        public static string PATH_INPUT_EXCEL = GetValue("files", "pathExcelInput");
        public static string PATH_OUTPUT_EXCEL = GetValue("files", "pathExcelOutput");
        public static string FILE_EXTENSION = GetValue("files", "fileExtension");
        public static string GetValue(string section, string key)
        {
            var iniFile = new IniFile(@"C:\Users\night\Desktop\RPA Challenge Csharp\rpa.challenge\rpa.challenge\Resources\config.ini");
            return iniFile.GetValue(section, key);
        }
    }
}
