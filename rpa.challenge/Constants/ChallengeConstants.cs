using OpenQA.Selenium;
using rpa.challenge.utils;


namespace rpa.challenge.Constants
{
    class ChallengeConstants
    {
        public static string URL_CHALLENGE = GetValue("url", "mainUrlChallenge");
        public static string PATH_INPUT_EXCEL = GetValue("files", "pathExcelInput");
        public static string PATH_OUTPUT_EXCEL = GetValue("files", "pathExcelOutput");
        public static string FILE_EXTENSION = GetValue("files", "fileExtension");

        public static By XPATH_START_BUTTON = By.XPath("//button[text()='Start']");
        public static By XPATH_SUBMIT_BUTTON = By.XPath("//input[@value = 'Submit']");
        public static string XPATH_FINAL_MESSAGE = "//div[contains(@class,'message2')]";

        public static By byXpathFirstName = By.XPath("//label[text()='First Name']/following-sibling::input");
        public static By byXpathLastName = By.XPath("//label[text()='Last Name']/following-sibling::input");
        public static By byXpathCompanyName = By.XPath("//label[text()='Company Name']/following-sibling::input");
        public static By byXpathPhoneNumber = By.XPath("//label[text()='Phone Number']/following-sibling::input");
        public static By byXpathEmail = By.XPath("//label[text()='Email']/following-sibling::input");
        public static By byXpathRole = By.XPath("//label[text()='Role in Company']/following-sibling::input");
        public static By byXpathAddress = By.XPath("//label[text()='Address']/following-sibling::input");

        public static string GetValue(string section, string key)
        {
            var iniFile = new IniFile(@"C:\Users\night\Desktop\RPA Challenge Csharp\rpa.challenge\rpa.challenge\Resources\config.ini");
            return iniFile.GetValue(section, key);
        }
    }
}
