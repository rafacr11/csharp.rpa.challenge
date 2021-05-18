using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using rpa.challenge.Model;
using System;
using System.Collections.Generic;
using System.IO;
using rpa.challenge.Controller;

namespace rpa.challenge
{
    class Program
    {
        static void Main(string[] args)
        {              
            RpaController controller = new RpaController();
            controller.StartRobot();
        }
    }
}
