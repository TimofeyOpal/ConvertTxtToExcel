using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Reflection;
using OfficeOpenXml;

namespace parceText
{
    class Program
    {
        static void Main(string[] args)
        {
            Converter();


        }

        public static void Converter()
        {
            Console.WriteLine("Введите путь до папки с файлами в формате 'D:\\folder with files'");
            Console.Write("Ваш путь:  ");
            string pathToFolderWithFiles = Console.ReadLine().Trim();
            List<List<string>> list = new();
            string[] fileEntries = Directory.GetFiles(pathToFolderWithFiles);
            var a = -1;
            var step = -5;
            foreach (var item in fileEntries)
            {
                a++;
                string[] readText = File.ReadAllLines(item);
                step += 5;
                for (int i = 0; i < readText.Length; i++)
                {
                    if (a == 0)
                    {
                        list.Add(new List<string>(readText[i].Split('\t')));
                    }
                    else
                    {
                        try
                        {
                            list[i].InsertRange(step, new List<string>(readText[i].Split('\t')));
                        }
                        catch
                        {
                            List<string> whiteList = new();
                            for (int c = 0; c < step; c++)
                                whiteList.Add("");
                            list.Add(new List<string>(whiteList));
                            list[i].InsertRange(step, new List<string>(readText[i].Split('\t')));
                        }
                    }
                }
                a = 0;
            }


            Console.WriteLine("Работа завершена. Введите путь для сохранения вашего файла в формате D:\\MyWorkbook.xlsx");
            Console.WriteLine("Такой формат не доступен, при сохранении на диск С указываете дальнейший путь пример: С:\\my folder\\MyWorkbook.xlsx");
            Console.Write("Ввод");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Program.CreateExcelFile(Console.ReadLine().Trim(), list);
        }

        public static void CreateExcelFile(string fileName, List<List<string>> list)
        {
            if (File.Exists(fileName)) File.Delete(fileName);
            using (var excel = new ExcelPackage(new FileInfo(fileName)))
            {
                
                int j = 0;
                var ws = excel.Workbook.Worksheets.Add("Sheet");
                
                for (int i = 0; i < list.Count; i++)
                {
                    foreach (var item in list[i])
                    {
                        j++;
                        ws.Cells[i+1, j].Value = item;
                    }
                    j = 0;
                    
                }
                excel.Save();
            }
        }

    }
}
