using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReplaceListaDeEmails
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string inputFilePath = "D:\\Usuários\\Cliente\\Desktop\\bounce_simpar.xlsx";
            string outputFilePath = "D:\\Usuários\\Cliente\\Documents\\output.xlsx";

            List<string> items = ReadFromExcel(inputFilePath);
            List<string> modifiedItems = AddQuotesToList(items);
            WriteToExcel(outputFilePath, modifiedItems);

            Console.WriteLine("Process completed successfully.");
        }

        static List<string> ReadFromExcel(string filePath)
        {
            List<string> items = new List<string>();

            FileInfo fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    string item = worksheet.Cells[row, 1].Text;
                    items.Add(item);
                }
            }

            return items;
        }
        static List<string> AddQuotesToList(List<string> items)
        {
            List<string> modifiedItems = new List<string>();

            foreach (string item in items)
            {
                string modifiedItem = $"'{item}',";
                modifiedItems.Add(modifiedItem);
            }

            return modifiedItems;
        }

        static void WriteToExcel(string filePath, List<string> items)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("ModifiedItems");

                for (int row = 0; row < items.Count; row++)
                {
                    worksheet.Cells[row + 1, 1].Value = items[row];
                }

                package.Save();
            }
        }
    }
}

