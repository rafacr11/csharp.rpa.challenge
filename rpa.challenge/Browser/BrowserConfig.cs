using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rpa.challenge.Browser
{
    class BrowserConfig
    {
        public static ChromeOptions GetChromeOptions()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--start-maximized");
            options.AddArguments("--disable-extensions");
            options.AddArguments("--disable-default-apps");
            options.AddArguments("--disable-infobars");

            return options;
        }
    }
}
