using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using rpa.challenge.Constants;
using rpa.challenge.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rpa.challenge.Controller
{
    class RpaController
    {
        private static IWebDriver chromeDriver;
        public void StartRobot()
        {
            ExcelController excelController = new ExcelController();
            List<Pessoa> pessoas = excelController.ReadExcel();

            if (pessoas.Count != 0)
            {
                NavigateRpaChallenge(pessoas);
                excelController.OutputExcelFile(pessoas);
            }
            else
            {
                Console.WriteLine("Nenhum registro para processar");
            }  
        }

        private void NavigateRpaChallenge(List<Pessoa> pessoas)
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--start-maximized");
            options.AddArguments("--disable-extensions");
            options.AddArguments("--disable-default-apps");
            options.AddArguments("--disable-infobars");

            chromeDriver = new ChromeDriver(options);

           
            

            chromeDriver.Navigate().GoToUrl(ChallengeConstants.URL_CHALLENGE);
            chromeDriver.FindElement(By.XPath("//button[text()='Start']")).Click();

            foreach (Pessoa pessoa in pessoas)
            {
                pessoa.isProcessed = InsertData(pessoa);
            }

            string finalmessage = chromeDriver.FindElement(By.XPath("//div[contains(@class,'message2')]")).Text;
            Console.WriteLine("");
            Console.WriteLine("---------------------------------------------");
            Console.WriteLine(finalmessage);
            
            chromeDriver.Close();
            chromeDriver.Quit();
        }

        private bool InsertData(Pessoa pessoa)
        {
            bool isProcessed = false;
            try
            {
                chromeDriver.FindElement(By.XPath("//label[text()='First Name']/following-sibling::input")).SendKeys(pessoa.firstname);
                chromeDriver.FindElement(By.XPath("//label[text()='Company Name']/following-sibling::input")).SendKeys(pessoa.companyname);
                chromeDriver.FindElement(By.XPath("//label[text()='Phone Number']/following-sibling::input")).SendKeys(pessoa.phonenumber);
                chromeDriver.FindElement(By.XPath("//label[text()='Email']/following-sibling::input")).SendKeys(pessoa.email);
                chromeDriver.FindElement(By.XPath("//label[text()='Last Name']/following-sibling::input")).SendKeys(pessoa.lastname);
                chromeDriver.FindElement(By.XPath("//label[text()='Role in Company']/following-sibling::input")).SendKeys(pessoa.role);
                chromeDriver.FindElement(By.XPath("//label[text()='Address']/following-sibling::input")).SendKeys(pessoa.adress);
                chromeDriver.FindElement(By.XPath("//input[@value = 'Submit']")).Click();

                isProcessed = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return isProcessed;
        }
    }
}
