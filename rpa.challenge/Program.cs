using System;
using rpa.challenge.Controller;
using rpa.challenge.utils;

namespace rpa.challenge
{
    class Program
    {
        private static readonly Logging log = new(@"C:\Users\night\Desktop\RPA Challenge Csharp\rpa.challenge\logs", "log_" + DateTime.Now.ToString("yyyMMdd_HHmmss") + ".txt");
        static void Main(string[] args)
        {
            log.Info("Iniciando aplicação");

            RpaController controller = new RpaController(log);
            controller.StartRobot();

            log.Info("Execução terminada");            
        }
    }
}
