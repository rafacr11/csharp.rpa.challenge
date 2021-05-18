using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
    class ExcelController
    {
        private string fileExtension = ChallengeConstants.FILE_EXTENSION;
        public List<Pessoa> ReadExcel()
        {
            string filePath = ChallengeConstants.PATH_INPUT_EXCEL + fileExtension;
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            IWorkbook book;
            book = new XSSFWorkbook(fs);

            ISheet sheet = book.GetSheetAt(0);

            List<Pessoa> pessoas = new List<Pessoa>();

            if (sheet != null)
            {
                // int rowCount = sheet.LastRowNum;
                for (int i = 1; i <= 10; i++)
                {
                    IRow curRow = sheet.GetRow(i);

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

            return pessoas;
        }
        public void OutputExcelFile(List<Pessoa> pessoas)
        {
            IWorkbook OutPutBook;
            OutPutBook = new XSSFWorkbook();

            ISheet outputsheet = OutPutBook.CreateSheet();

            int rowIndex = 1;

            IRow headersRow = outputsheet.CreateRow(0);
            ICell headerCell = headersRow.CreateCell(0);
            headerCell.SetCellValue("First Name");

            headerCell = headersRow.CreateCell(1);
            headerCell.SetCellValue("Last Name");

            headerCell = headersRow.CreateCell(2);
            headerCell.SetCellValue("Company Name");

            headerCell = headersRow.CreateCell(3);
            headerCell.SetCellValue("Role in Company");

            headerCell = headersRow.CreateCell(4);
            headerCell.SetCellValue("Address");

            headerCell = headersRow.CreateCell(5);
            headerCell.SetCellValue("Email");

            headerCell = headersRow.CreateCell(6);
            headerCell.SetCellValue("Phone Number");

            headerCell = headersRow.CreateCell(7);
            headerCell.SetCellValue("Status");

            foreach (Pessoa pessoa in pessoas)
            {
                IRow row = outputsheet.CreateRow(rowIndex++);

                ICell cell = row.CreateCell(0);
                cell.SetCellValue(pessoa.firstname);

                cell = row.CreateCell(1);
                cell.SetCellValue(pessoa.lastname);

                cell = row.CreateCell(2);
                cell.SetCellValue(pessoa.companyname);

                cell = row.CreateCell(3);
                cell.SetCellValue(pessoa.role);

                cell = row.CreateCell(4);
                cell.SetCellValue(pessoa.adress);

                cell = row.CreateCell(5);
                cell.SetCellValue(pessoa.email);

                cell = row.CreateCell(6);
                cell.SetCellValue(pessoa.phonenumber);

                cell = row.CreateCell(7);
                cell.SetCellValue(pessoa.isProcessed);
            }

            string targetPath = ChallengeConstants.PATH_OUTPUT_EXCEL + fileExtension;
            using (FileStream arquivo = File.Create(targetPath))
            {
                OutPutBook.Write(arquivo);
            }

        }

    }
}
