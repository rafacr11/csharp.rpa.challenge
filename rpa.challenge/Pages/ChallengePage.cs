using OpenQA.Selenium;
using rpa.challenge.Constants;
using rpa.challenge.Model;
using rpa.challenge.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rpa.challenge.Pages
{
    class ChallengePage
    {
        private IWebDriver driver;
        private readonly Logging log;
        public ChallengePage(IWebDriver driver,Logging log)
        {
            this.driver = driver;
            this.log = log;
        }
       
        public void navigateChallengeUrl()
        {
            driver.Navigate().GoToUrl(ChallengeConstants.URL_CHALLENGE);
        }

        public void clickStart()
        {
            driver.FindElement(ChallengeConstants.XPATH_START_BUTTON).Click();
        }

        public void clickSubmit()
        {
            driver.FindElement(ChallengeConstants.XPATH_SUBMIT_BUTTON).Click();
        }
       
        public bool InsertData(Pessoa pessoa)
        {
            bool isProcessed = false;
            try
            {
                driver.FindElement(ChallengeConstants.byXpathFirstName).SendKeys(pessoa.firstname);
                driver.FindElement(ChallengeConstants.byXpathCompanyName).SendKeys(pessoa.companyname);
                driver.FindElement(ChallengeConstants.byXpathPhoneNumber).SendKeys(pessoa.phonenumber);
                driver.FindElement(ChallengeConstants.byXpathEmail).SendKeys(pessoa.email);
                driver.FindElement(ChallengeConstants.byXpathLastName).SendKeys(pessoa.lastname);
                driver.FindElement(ChallengeConstants.byXpathRole).SendKeys(pessoa.role);
                driver.FindElement(ChallengeConstants.byXpathAddress).SendKeys(pessoa.adress);
                clickSubmit();      
                
                isProcessed = true;
                log.Info(pessoa.firstname + " inserida com sucesso! ");
            }
            catch (Exception e)
            {
                driver.FindElement(ChallengeConstants.byXpathFirstName).Clear();
                driver.FindElement(ChallengeConstants.byXpathCompanyName).Clear();
                driver.FindElement(ChallengeConstants.byXpathPhoneNumber).Clear();
                driver.FindElement(ChallengeConstants.byXpathEmail).Clear();
                driver.FindElement(ChallengeConstants.byXpathLastName).Clear();
                driver.FindElement(ChallengeConstants.byXpathRole).Clear();
                driver.FindElement(ChallengeConstants.byXpathAddress).Clear();
                isProcessed = false;

                log.Info("Erro ao inserir: " + pessoa.firstname + " Erro: " + e.Message);
            }
            return isProcessed;
        }
        public void WriteFinalMessage()
        {
            string finalMessage = driver.FindElement(By.XPath("//div[contains(@class,'message2')]")).Text;
            Console.WriteLine("");
            Console.WriteLine("---------------------------------------------");
            Console.WriteLine(finalMessage);
        }
    }
}
