using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using rpa.challenge.Browser;
using rpa.challenge.Constants;
using rpa.challenge.Model;
using rpa.challenge.Pages;
using rpa.challenge.utils;
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
        private readonly Logging log;
        private static IWebDriver chromeDriver;
        ChallengePage challengePage;

        public RpaController(Logging log)
        {
            this.log = log;
        }

        public void StartRobot()
        {
            ExcelController excelController = new ExcelController(log);
            List<Pessoa> pessoas = excelController.ReadExcel();

            if (pessoas.Count != 0)
            {
                log.Info("Quantidade de registros na planilha: " + pessoas.Count());
                NavigateRpaChallenge(pessoas);
                excelController.OutputExcelFile(pessoas);

                log.Info("Arquivos processados");
            }
            else
            {
                log.Info("Não existem arquivos para serem processados");
            }  
        }

        private void NavigateRpaChallenge(List<Pessoa> pessoas)
        {          
            chromeDriver = new ChromeDriver(BrowserConfig.GetChromeOptions());                
            this.challengePage = new ChallengePage(chromeDriver,log);
            challengePage.navigateChallengeUrl();
            challengePage.clickStart();

            foreach (Pessoa pessoa in pessoas)
            {                
                pessoa.isProcessed = challengePage.InsertData(pessoa);              
            }

            challengePage.WriteFinalMessage();
          
            chromeDriver.Close();
            chromeDriver.Quit();
        }  
    }
}
