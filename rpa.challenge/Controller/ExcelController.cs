using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using rpa.challenge.Constants;
using rpa.challenge.Model;
using rpa.challenge.utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rpa.challenge.Controller
{
    class ExcelController
    {
        private string fileExtension = ChallengeConstants.FILE_EXTENSION;
        private readonly Logging log;

        public ExcelController(Logging log)
        {
            this.log = log;
        }
        public List<Pessoa> ReadExcel()
        {
            List<Pessoa> pessoas = new List<Pessoa>();
            string filePath = ChallengeConstants.PATH_INPUT_EXCEL + fileExtension;
            try
            {                
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                IWorkbook book;
                book = new XSSFWorkbook(fs);

                ISheet sheet = book.GetSheetAt(0);                


                foreach(IRow curRow in sheet)
                {
                    String firstCell = curRow.GetCell(0).StringCellValue;
                    if (!firstCell.Equals("") && !firstCell.Equals("First Name")){
                        sheet.GetRow(curRow.RowNum);

                        Pessoa pessoa = new Pessoa();

                        pessoa.firstname = curRow.GetCell(0).StringCellValue.Trim();
                        pessoa.lastname = curRow.GetCell(1).StringCellValue.Trim();
                        pessoa.companyname = curRow.GetCell(2).StringCellValue.Trim();
                        pessoa.role = curRow.GetCell(3).StringCellValue.Trim();
                        pessoa.adress = curRow.GetCell(4).StringCellValue.Trim();
                        pessoa.email = curRow.GetCell(5).StringCellValue.Trim();
                        pessoa.phonenumber = curRow.GetCell(6).NumericCellValue.ToString();

                        pessoas.Add(pessoa);
                    }
                }
            }catch (Exception e)
            {
                log.Error("Falha ao ler Excel: " + e.Message);
            }


            return pessoas;

        }

        public void OutputExcelFile(List<Pessoa> pessoas)
        {
            try {
                IWorkbook OutPutBook;
                OutPutBook = new XSSFWorkbook();

                ISheet outputsheet = OutPutBook.CreateSheet();

                int rowIndex = 1;

                IRow headersRow = outputsheet.CreateRow(0);
                headersRow.CreateCell(0).SetCellValue("First Name");
                headersRow.CreateCell(1).SetCellValue("Last Name");
                headersRow.CreateCell(2).SetCellValue("Company Name");
                headersRow.CreateCell(3).SetCellValue("Role in Company");         
                headersRow.CreateCell(4).SetCellValue("Address");
                headersRow.CreateCell(5).SetCellValue("Email");               
                headersRow.CreateCell(6).SetCellValue("Phone Number");            
                headersRow.CreateCell(7).SetCellValue("Status");             

                foreach (Pessoa pessoa in pessoas)
                {
                    IRow row = outputsheet.CreateRow(rowIndex++);
                    row.CreateCell(0).SetCellValue(pessoa.firstname);                    
                    row.CreateCell(1).SetCellValue(pessoa.lastname);
                    row.CreateCell(2).SetCellValue(pessoa.companyname);                    
                    row.CreateCell(3).SetCellValue(pessoa.role);
                    row.CreateCell(4).SetCellValue(pessoa.adress);
                    row.CreateCell(5).SetCellValue(pessoa.email);
                    row.CreateCell(6).SetCellValue(pessoa.phonenumber);               
                    row.CreateCell(7).SetCellValue(pessoa.isProcessed);
                }

                string targetPath = ChallengeConstants.PATH_OUTPUT_EXCEL + fileExtension;
                using (FileStream arquivo = File.Create(targetPath))
                {
                    OutPutBook.Write(arquivo);
                    OutPutBook.Close();                   
                }

                log.Info("Arquivo exportado com sucesso !");
            } catch (Exception e)
            { 
                log.Error("Erro ao exportar arquivo"  + e.Message);
            }
            

        }

    }
}
